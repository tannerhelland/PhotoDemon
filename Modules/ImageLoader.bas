Attribute VB_Name = "ImageImporter"
'***************************************************************************
'Low-level image import interfaces
'Copyright 2001-2016 by Tanner Helland
'Created: 4/15/01
'Last updated: 09/March/16
'Last update: migrate various functions out of the high-level "Loading" module and into this new, format-specific module
'
'This module provides low-level "import" functionality for importing image files into PD.  You will not generally want
' to interface with this module directly; instead, rely on the high-level functions in the "Loading" module.
' They will intelligently drop into this module as necessary, sparing you the messy work of having to handle
' format-specific details (which are many).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_JpegObeyEXIFOrientation As PD_BOOL

'Some user preferences control how image importing behaves.  Because these preferences are accessed frequently, we cache them
' locally improve performance.  External functions should use our wrappers instead of accessing the preferences directly.
' Also, changes to these preferences obviously require a re-cache; use the reset function, below, for that.
Public Sub ResetImageImportPreferenceCache()
    m_JpegObeyEXIFOrientation = PD_BOOL_UNKNOWN
End Sub

Public Function GetImportPref_JPEGOrientation() As Boolean
    If (m_JpegObeyEXIFOrientation = PD_BOOL_UNKNOWN) Then
        If g_UserPreferences.GetPref_Boolean("Loading", "ExifAutoRotate", True) Then m_JpegObeyEXIFOrientation = PD_BOOL_TRUE Else m_JpegObeyEXIFOrientation = PD_BOOL_FALSE
    End If
    GetImportPref_JPEGOrientation = CBool(m_JpegObeyEXIFOrientation = PD_BOOL_TRUE)
End Function

'PDI loading.  "PhotoDemon Image" files are the only format PD supports for saving layered images.  PDI to PhotoDemon is like
' PSD to PhotoShop, or XCF to Gimp.
'
'Note the unique "sourceIsUndoFile" parameter for this load function.  PDI files are used to store undo/redo data, and when one of their
' kind is loaded as part of an Undo/Redo action, we must ignore certain elements stored in the file (e.g. settings like "LastSaveFormat"
' which we do not want to Undo/Redo).  This parameter is passed to the pdImage initializer, and it tells it to ignore certain settings.
Public Function LoadPhotoDemonImage(ByVal pdiPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "PDI file identified.  Starting pdPackage decompression..."
        Dim startTime As Currency
        VB_Hacks.GetHighResTime startTime
    #End If
    
    On Error GoTo LoadPDIFail
    
    'PDI files require a parent pdImage container
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager2
    Set pdiReader = New pdPackager2
    
    'Load the file into the pdPackager instance.  Note that this step will also validate the incoming file.
    ' (Also, prior to v7.0, PD would copy the entire source file into memory, then load the PDI from there.  This no longer occurs;
    '  instead, the file is left on-disk, and data is only loaded on a per-node basis.  This greatly reduces memory load.)
    ' (Also, because PDI files store data roughly sequentially, we can use OptimizeSequentialAccess for a small perf boost.)
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER, PD_SM_MemoryBacked, PD_SA_ReadOnly, OptimizeSequentialAccess) Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "pdPackage successfully read and initialized.  Starting package parsing..."
        #End If
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String, retSize As Long
        
        If pdiReader.GetNodeDataByIndex(0, True, retBytes, False, retSize) Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Initial PDI node retrieved.  Initializing corresponding pdImage object..."
            #End If
            
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.ReadExternalData retString, True, sourceIsUndoFile
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "pdImage created successfully.  Moving on to individual layers..."
        #End If
        
        'With the main pdImage now assembled, the next task is to populate all layers with two pieces of information:
        ' 1) The layer header, which contains stuff like layer name, opacity, blend mode, etc
        ' 2) Layer-specific information, which varies by layer type.  For DIBs, this will be a raw stream of bytes
        '    containing the layer DIB's raster data.  For text or other vector layers, this is an XML stream containing
        '    whatever information is necessary to construct the layer from scratch.
        
        Dim i As Long
        For i = 0 To dstImage.GetNumOfLayers - 1
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Retrieving layer header " & i & "..."
            #End If
        
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
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Raster layer identified.  Retrieving pixel bits..."
                #End If
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstImage.GetLayerByIndex(i).layerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                'At present, all pdPackage layers will contain premultiplied alpha, so force the corresponding state now
                dstImage.GetLayerByIndex(i).layerDIB.SetInitialAlphaPremultiplicationState True
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByIndex_UnsafeDstPointer(i + 1, False, tmpDIBPointer)
            
            'Text and other vector layers
            ElseIf dstImage.GetLayerByIndex(i).IsLayerVector Then
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Vector layer identified.  Retrieving layer XML..."
                #End If
                
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
                dstImage.NotifyImageChanged UNDO_LAYER, i
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        Next i
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "All layers loaded.  Looking for remaining non-essential PDI data..."
        #End If
        
        'Finally, check to see if the PDI image has a metadata entry.  If it does, load that data now.
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", True, retBytes, False, retSize) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Raw metadata chunk found.  Retrieving now..."
            #End If
        
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            If Not dstImage.imgMetadata.LoadAllMetadata(retString, dstImage.imageID) Then
                
                'For invalid metadata, do not reject the rest of the PDI file.  Instead, just warn the user and carry on.
                Debug.Print "PDI Metadata Node rejected by metadata parser."
                
            End If
        
        End If
        
        '(As of v7.0, a serialized copy of the image's metadata is also stored.  This copy contains all user edits
        ' and other changes.)
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", False, retBytes, False, retSize) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Serialized metadata chunk found.  Retrieving now..."
            #End If
        
            'Copy the received bytes into a string
            retString = Space$(retSize \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            dstImage.imgMetadata.RecreateFromSerializedXMLData retString
        
        End If
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "PDI parsing complete.  Returning control to main image loader..."
            Debug.Print "Time required to load PDI file: " & Format$(VB_Hacks.GetTimerDifferenceNow(startTime) * 1000, "####0.00") & " ms"
        #End If
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank 16, 16, 32, 0
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImage = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed.  This may be a legacy PDI file -- try that function next.
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Legacy PDI file encountered; dropping back to pdPackage v1 functions..."
        #End If
        LoadPhotoDemonImage = LoadPDI_Legacy(pdiPath, dstDIB, dstImage, sourceIsUndoFile)
    
    End If
    
    Exit Function
    
LoadPDIFail:
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    #End If
    
    'Before falling back to a generic error message, check for a couple known problem states.
    
    'Case 1: zLib is required for this file, but the user doesn't have the zLib plugin
    If pdiReader.GetPackageFlag(PDP_FLAG_ZLIB_REQUIRED, PDP_LOCATION_ANY) And (Not g_ZLibEnabled) Then
        PDMsgBox "The PDI file ""%1"" contains compressed data, but the zLib plugin is missing or disabled." & vbCrLf & vbCrLf & "To enable support for compressed PDI files, click Help > Check for Updates, and when prompted, allow PhotoDemon to download all recommended plugins.", vbInformation + vbOKOnly + vbApplicationModal, "zLib plugin missing", GetFilename(pdiPath)
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
    Dim pdiReader As pdPackager2
    Set pdiReader = New pdPackager2
    
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
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImageHeaderOnly = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file may be a legacy PDI format.
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "PDI v2 validation failed.  Attempting v1 load engine..."
        #End If
        
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
    Dim pdiReader As pdPackager2
    Set pdiReader = New pdPackager2
    
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
                    ' as vector layers were implemented after pdPackager was given Unicode compatibility.
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
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "PDI v2 validation failed.  Attempting v1 load engine..."
        #End If
        
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
    Dim pdiReader As pdPackager2
    Set pdiReader = New pdPackager2
    
    'Load the file into the pdPackager instance.  pdPackager It will cache the file contents, so we only have to do this once.
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
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager was given Unicode compatibility.
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
Public Function LoadGDIPlusImage(ByVal imagePath As String, ByRef dstDIB As pdDIB) As Boolean
    Dim verifyGDISuccess As Boolean
    verifyGDISuccess = GDIPlusLoadPicture(imagePath, dstDIB)
    If verifyGDISuccess Then
        If (Not dstDIB Is Nothing) Then
            LoadGDIPlusImage = CBool((dstDIB.GetDIBWidth <> 0) And (dstDIB.GetDIBHeight <> 0))
        Else
            LoadGDIPlusImage = False
        End If
    Else
        LoadGDIPlusImage = False
    End If
End Function

'BITMAP loading
Public Function LoadVBImage(ByVal imagePath As String, ByRef dstDIB As pdDIB) As Boolean
    
    On Error GoTo LoadVBImageFail
    
    'Create a temporary StdPicture object that will be used to load the image
    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
    Set tmpPicture = LoadPicture(imagePath)
    
    If ((tmpPicture.Width = 0) Or (tmpPicture.Height = 0)) Then
        LoadVBImage = False
        Exit Function
    End If
    
    'Copy the image into the current pdImage object
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateFromPicture tmpPicture
    
    LoadVBImage = True
    Exit Function
    
LoadVBImageFail:

    LoadVBImage = False
    Exit Function
    
End Function

Public Function IsFileSVGCandidate(ByVal imagePath As String) As Boolean
    IsFileSVGCandidate = CBool(StrComp(LCase$(Right$(imagePath, 3)), "svg", vbBinaryCompare) = 0)
    
    'Compressed SVG files are not currently supported.  (For them to work, we'd need to decompress to a temp file, which causes
    ' some messy interaction details with ExifTool - we'll deal with this in the future.)
    'If (Not IsFileSVGCandidate) Then IsFileSVGCandidate = CBool(StrComp(LCase$(Right$(imagePath, 4)), "svgz", vbBinaryCompare) = 0)
End Function

'SVG support is *experimental only*!  This function should not be enabled in production builds.
Public Function LoadSVG(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As Boolean
    
    On Error GoTo LoadSVGFail
    
    'In the future, we'll add meaningful heuristics, but for now, don't even attempt a load unless the file extension matches.
    If IsFileSVGCandidate(imagePath) Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Waiting for SVG parsing to complete..."
        #End If
        
        'Hang out while we wait for ExifTool to finish processing this image's metadata
        Do While (Not dstImage.imgMetadata.HasMetadata)
            DoEvents
            If ExifTool.IsMetadataFinished Then
                dstImage.imgMetadata.LoadAllMetadata ExifTool.RetrieveMetadataString, dstImage.imageID
            End If
        Loop
        
        'Retrieve the target SVG's width and height
        Dim svgWidth As String, svgHeight As String
        svgWidth = dstImage.imgMetadata.GetTagValue("SVG:ImageWidth", vbBinaryCompare, True)
        svgHeight = dstImage.imgMetadata.GetTagValue("SVG:ImageHeight", vbBinaryCompare, True)
        
        'If there's a viewbox, grab it too
        Dim svgHasViewbox As Boolean
        svgHasViewbox = dstImage.imgMetadata.DoesTagExistFullName("SVG:ViewBox", , vbBinaryCompare)
        
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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  SVG parsing failed with error #" & Err.Number & ": " & Err.Description
    #End If
    
    LoadSVG = False
    Exit Function
    
End Function

'Load data from a PD-generated Undo file.  This function is fairly complex, on account of PD's new diff-based Undo engine.
' Note that two types of Undo data must be specified: the Undo type of the file requested (because this function has no
' knowledge of that, by design), and what type of Undo data the caller wants extracted from the file.
'
'New as of 11 July '14 is the ability to specify a custom layer destination, for layer-relevant load operations.  If this value is NOTHING,
' the function will automatically load the data to the relevant layer in the parent pdImage object.  If this layer is supplied, however,
' the supplied layer reference will be used instead.
Public Sub LoadUndo(ByVal undoFile As String, ByVal undoTypeOfFile As Long, ByVal undoTypeOfAction As Long, Optional ByVal targetLayerID As Long = -1, Optional ByVal suspendRedraw As Boolean = False, Optional ByRef customLayerDestination As pdLayer = Nothing)
    
    'Certain load functions require access to a DIB, so declare a generic one in advance
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If selection data was loaded as part of this diff, this value will be set to TRUE.  We check it at the end of
    ' the load function, and activate various selection-related items as necessary.
    Dim selectionDataLoaded As Boolean
    selectionDataLoaded = False
    
    'Depending on the Undo data requested, we may end up loading one or more diff files at this location
    Select Case undoTypeOfAction
    
        'UNDO_EVERYTHING: a full copy of both the pdImage stack and all selection data is wanted
        Case UNDO_EVERYTHING
            ImageImporter.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            pdImages(g_CurrentImage).mainSelection.ReadSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
        'UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE: a full copy of the pdImage stack is wanted
        '             Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_IMAGE/_VECTORSAFE, we
        '             don't have to do any special processing to the file - just load the whole damn thing.
        Case UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE
            ImageImporter.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            
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
        Case UNDO_IMAGEHEADER
            ImageImporter.LoadPhotoDemonImageHeaderOnly undoFile, pdImages(g_CurrentImage)
            
            'Once the full image has been loaded, we now know that at least the *existence* of all layers is correct.
            ' Unfortunately, subsequent changes to the pdImage header (or individual layers/layer headers) still need
            ' to be manually reconstructed, because they may have changed between the last full pdImage write and the
            ' current image state.  This step is handled by the Undo/Redo engine, which will call this LoadUndo function
            ' as many times as necessary to reconstruct each individual layer against its most recent diff.
        
        'UNDO_LAYER, UNDO_LAYER_VECTORSAFE: a full copy of the saved layer data at this position.
        '             Because the underlying file data can be different types (layer data can be loaded from standalone layer saves,
        '             or from a full pdImage stack save), we must check the undo type of the saved file, and modify our load
        '             behavior accordingly.
        Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
            
            'New as of 11 July '14 is the ability for the caller to supply their own destination layer for layer-specific Undo data.
            ' Check this optional parameter, and if it is NOT supplied, point it at the relevant layer in the parent pdImage object.
            If (customLayerDestination Is Nothing) Then Set customLayerDestination = pdImages(g_CurrentImage).GetLayerByID(targetLayerID)
            
            'Layer data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer with the data from the file.
                Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
                    ImageImporter.LoadPhotoDemonLayer undoFile & ".layer", customLayerDestination, False
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_EVERYTHING, UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE
                    ImageImporter.LoadSingleLayerFromPDI undoFile, customLayerDestination, targetLayerID, False
                
            End Select
        
        'UNDO_LAYERHEADER: a full copy of the saved layer header data at this position.  Layer DIB data is ignored.
        '             Because the underlying file data can be many different types (layer data header can be loaded from
        '             standalone layer header saves, or full layer saves, or even a full pdImage stack), we must check the
        '             undo type of the saved file, and modify our load behavior accordingly.
        Case UNDO_LAYERHEADER
            
            'Layer header data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer header with the
                ' header data from this file.
                Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE, UNDO_LAYERHEADER
                    ImageImporter.LoadPhotoDemonLayer undoFile & ".layer", pdImages(g_CurrentImage).GetLayerByID(targetLayerID), True
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_EVERYTHING, UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE, UNDO_IMAGEHEADER
                    ImageImporter.LoadSingleLayerFromPDI undoFile, pdImages(g_CurrentImage).GetLayerByID(targetLayerID), targetLayerID, True
                
            End Select
        
        'UNDO_SELECTION: a full copy of the saved selection data is wanted
        '                 Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_SELECTION, we don't have to do
        '                 any special processing.
        Case UNDO_SELECTION
            pdImages(g_CurrentImage).mainSelection.ReadSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
            
        'For now, any unhandled Undo types result in a request for the full pdImage stack.  This line can be removed when
        ' all Undo types finally have their own custom handling implemented.
        Case Else
            ImageImporter.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            
        
    End Select
    
    'If a selection was loaded, activate all selection-related stuff now
    If selectionDataLoaded Then
    
        'Activate the selection as necessary
        pdImages(g_CurrentImage).selectionActive = pdImages(g_CurrentImage).mainSelection.IsLockedIn
        
        'Synchronize the text boxes as necessary
        syncTextToCurrentSelection g_CurrentImage
    
    End If
    
    'If a selection is active, request a redraw of the selection mask before rendering the image to the screen.  (If we are
    ' "undoing" an action that changed the image's size, the selection mask will be out of date.  Thus we need to re-render
    ' it before rendering the image or OOB errors may occur.)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.RequestNewMask
        
    'Render the image to the screen, if requested
    If Not suspendRedraw Then Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
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
Public Function CascadeLoadGenericImage(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_IMAGE_DECODER_ENGINE, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long) As Boolean
    
    CascadeLoadGenericImage = False
    
    'Before jumping out to a 3rd-party library, check for any image formats that we must decode using internal plugins.
    #If DEBUGMODE = 1 Then
        
        'SVG support is just experimental at present!
        CascadeLoadGenericImage = ImageImporter.LoadSVG(srcFile, dstDIB, dstImage)
        If CascadeLoadGenericImage Then
            decoderUsed = PDIDE_SVGPARSER
            dstImage.originalFileFormat = PDIF_SVG
            dstImage.SetDPI 96, 96
            dstImage.originalColorDepth = 32
        End If
        
    #End If
    
    'Note that FreeImage may raise additional dialogs (e.g. for HDR/RAW images), so it does not return a binary pass/fail.
    ' If the function fails due to user cancellation, we will suppress subsequent error message boxes.
    freeImage_Return = PD_FAILURE_GENERIC
    
    'If FreeImage is available, we first use it to try and load the image.
    If (Not CascadeLoadGenericImage) And g_ImageFormats.FreeImageEnabled Then
    
        'Start by seeing if the image file contains multiple pages.  If it does, we will load each page as a separate layer.
        ' TODO: preferences or prompt for how to handle such files??
        numOfPages = Plugin_FreeImage.IsMultiImage(srcFile)
        imageHasMultiplePages = (numOfPages > 1)
        freeImage_Return = FI_LoadImage_V5(srcFile, dstDIB)
        CascadeLoadGenericImage = CBool(freeImage_Return = PD_SUCCESS)
        
        'FreeImage worked!  Copy any relevant information from the DIB to the parent pdImage object (such as file format),
        ' then continue with the load process.
        If CascadeLoadGenericImage Then
            
            decoderUsed = PDIDE_FREEIMAGE
            
            dstImage.originalFileFormat = dstDIB.GetOriginalFormat
            dstImage.SetDPI dstDIB.GetDPI, dstDIB.GetDPI
            dstImage.originalColorDepth = dstDIB.GetOriginalColorDepth
            
            If (dstImage.originalFileFormat = PDIF_PNG) And (dstDIB.GetBackgroundColor <> -1) Then
                dstImage.imgStorage.AddEntry "pngBackgroundColor", dstDIB.GetBackgroundColor
            End If
            
        End If
        
    End If
            
    'If FreeImage fails for some reason, let GDI+ have a go at it.
    If (Not CascadeLoadGenericImage) And (freeImage_Return <> PD_FAILURE_USER_CANCELED) Then
        
        If g_ImageFormats.GDIPlusEnabled Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "FreeImage refused to load image.  Dropping back to GDI+ and trying again..."
            #End If
            
            CascadeLoadGenericImage = LoadGDIPlusImage(srcFile, dstDIB)
            
            If CascadeLoadGenericImage Then
                decoderUsed = PDIDE_GDIPLUS
                dstImage.originalFileFormat = dstDIB.GetOriginalFormat
                dstImage.SetDPI dstDIB.GetDPI, dstDIB.GetDPI
                dstImage.originalColorDepth = dstDIB.GetOriginalColorDepth
            End If
                
        End If
        
        'If GDI+ failed, run one last-ditch attempt using the classic OLE decoder.  (Note that we don't allow the OLE decoder to
        ' touch WMF or EMF files, as malformed ones can cause a silent API failure, bringing down the entire program.)
        If (Not CascadeLoadGenericImage) Then
            
            Dim srcFileExtension As String
            srcFileExtension = UCase(GetExtension(srcFile))
            
            If ((srcFileExtension <> "EMF") And (srcFileExtension <> "WMF")) Then
                #If DEBUGMODE = 1 Then
                    Message "GDI+ refused to load image.  Dropping back to internal routines and trying again..."
                #End If
                
                If LoadVBImage(srcFile, dstDIB) Then
                    CascadeLoadGenericImage = True
                    decoderUsed = PDIDE_VBLOADPICTURE
                    EstimateMissingMetadata dstImage, srcFileExtension
                End If
            End If
        End If
        
    End If
    
End Function

'Test an incoming image file against PD's internal decoder engines.  This function is much faster than
' CascadeLoadGenericImage(), above, and it should be preferentially used for image files generated by PD itself.
Public Function CascadeLoadInternalImage(ByVal internalFormatID As Long, ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, ByRef freeImage_Return As PD_OPERATION_OUTCOME, ByRef decoderUsed As PD_IMAGE_DECODER_ENGINE, ByRef imageHasMultiplePages As Boolean, ByRef numOfPages As Long) As Boolean
    
    Select Case internalFormatID
        
        Case PDIF_PDI
        
            'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
            CascadeLoadInternalImage = LoadPhotoDemonImage(srcFile, dstDIB, dstImage)
            
            dstImage.originalFileFormat = PDIF_PDI
            dstImage.originalColorDepth = 32
            dstImage.NotifyImageChanged UNDO_EVERYTHING
            decoderUsed = PDIDE_INTERNAL
            
        'TMPDIB files are raw pdDIB objects dumped directly to file.  In some cases, this is faster and easier for PD than wrapping
        ' the pdDIB object inside a pdPackage layer (e.g. during clipboard interactions, since we start with a raw pdDIB object
        ' after selections and such are applied to the base layer/image, so we may as well just use the raw pdDIB data we've cached).
        Case PDIF_RAWBUFFER
            
            'These raw pdDIB objects may require zLib for parsing (compression is optional), so it is possible for the load function
            ' to fail if zLib goes missing.
            CascadeLoadInternalImage = LoadRawImageBuffer(srcFile, dstDIB, dstImage)
            
            dstImage.originalFileFormat = PDIF_UNKNOWN
            dstImage.originalColorDepth = 32
            dstImage.NotifyImageChanged UNDO_EVERYTHING
            decoderUsed = PDIDE_INTERNAL
            
        'Straight TMP files are internal files (BMP, typically) used by PhotoDemon.  As ridiculous as it sounds, we must
        ' default to the generic load engine list, as the format of a TMP file is not guaranteed in advance.  Because of this,
        ' we can rely on the generic load engine to properly set things like "original color depth".
        '
        '(TODO: settle on a single tmp file format, so we don't have to play this game??)
        Case PDIF_TMPFILE
            CascadeLoadInternalImage = ImageImporter.CascadeLoadGenericImage(srcFile, dstImage, dstDIB, freeImage_Return, decoderUsed, imageHasMultiplePages, numOfPages)
            dstImage.originalFileFormat = PDIF_UNKNOWN
            
    End Select
    
End Function

'If a horribly broken image can't be loaded until we drop back to ancient OLE interfaces, we'll need to estimate
' or infer certain metadata bits (like original color-depth).  This method is only useful if FreeImage or GDI+
' was *not* used to load an image.
Private Sub EstimateMissingMetadata(ByRef dstImage As pdImage, ByRef srcFileExtension As String)
    
    Select Case srcFileExtension
                
        Case "GIF"
            dstImage.originalFileFormat = PDIF_GIF
            dstImage.originalColorDepth = 8
            
        Case "ICO"
            dstImage.originalFileFormat = FIF_ICO
        
        Case "JIF", "JFIF", "JPG", "JPEG", "JPE"
            dstImage.originalFileFormat = PDIF_JPEG
            dstImage.originalColorDepth = 24
            
        Case "PNG"
            dstImage.originalFileFormat = PDIF_PNG
        
        Case "TIF", "TIFF"
            dstImage.originalFileFormat = PDIF_TIFF
        
        Case "PDI", "TMP", "PDTMP", "TMPDIB", "PDTMPDIB"
            dstImage.originalFileFormat = PDIF_JPEG
            dstImage.originalColorDepth = 24
        
        'Treat anything else as a BMP file
        Case Else
            dstImage.originalFileFormat = PDIF_BMP
            
    End Select
    
End Sub

'See the Loading.LoadFileAsNewImage() function for where and when to apply this sub.
' IMPORTANT NOTE: some ICC profiles are applied to the image very early in the load process (e.g. CMYK, which requires a special pipeline).
'                 For such images, this function is meaningless.
'
'Returns: TRUE if changes were made to the target DIB
Public Function ApplyPostLoadICCHandling(ByRef targetDIB As pdDIB, Optional ByRef targetImage As pdImage = Nothing) As Boolean
    
    ApplyPostLoadICCHandling = False
    
    If targetDIB.ICCProfile.HasICCData Then
        If (Not targetDIB.ICCProfile.HasProfileBeenApplied) Then
        
            Dim colorManagementNeeded As Boolean
            If (targetImage Is Nothing) Then
                colorManagementNeeded = True
            Else
                colorManagementNeeded = (Not targetImage.imgStorage.DoesKeyExist("Tone-mapping"))
            End If
        
            If colorManagementNeeded Then
                
                If (targetDIB.GetDIBColorDepth = 32) Then targetDIB.SetAlphaPremultiplication False
                
                'During debug mode, color-management performance is an item of interest
                #If DEBUGMODE = 1 Then
                    Dim startTime As Currency
                    VB_Hacks.GetHighResTime startTime
                #End If
                
                'LittleCMS is our preferred color management engine.  Use it whenever possible.
                If g_LCMSEnabled Then
                    LittleCMS.ApplyICCProfileToPDDIB targetDIB
                Else
                    ColorManagement.ApplyICCtoPDDib_WindowsCMS targetDIB
                End If
                
                #If DEBUGMODE = 1 Then
                    Dim engineUsed As String
                    If g_LCMSEnabled Then engineUsed = "LittleCMS" Else engineUsed = "Windows ICM"
                    pdDebug.LogAction "Note: color management of the imported image took " & CStr(VB_Hacks.GetTimerDifferenceNow(startTime) * 1000) & " ms using " & engineUsed
                #End If
                
                If (targetDIB.GetDIBColorDepth = 32) Then targetDIB.SetAlphaPremultiplication True
                ApplyPostLoadICCHandling = True
                
            End If
            
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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "SyncRecoveredAutosaveImage invoked; attempting to recover usable data from the Autosave database..."
    #End If
    
    srcImage.imgStorage.AddEntry "CurrentLocationOnDisk", srcFile
            
    'Ask the AutoSave engine to synchronize this image's data against whatever it can recover from the Autosave database
    Autosave_Handler.AlignLoadedImageWithAutosave srcImage
            
    'This is a bit wacky, but - the Autosave engine will automatically update the "locationOnDisk" attribute based on
    ' information inside the Autosave recovery database.  We thus want to overwrite the original srcFile value (which points
    ' at a temp file copy of whatever we're attempting to recover), with the new, recovered srcFile value.
    srcFile = srcImage.imgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)
    
    SyncRecoveredAutosaveImage = True
    
End Function

'After loading an image file, you can call this function to set up any post-load pdImage attributes (like name and save state)
Public Function GenerateExtraPDImageAttributes(ByRef srcFile As String, ByRef targetImage As pdImage, ByRef suggestedFilename As String) As Boolean
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Determining initial pdImage attributes..."
    #End If
    
    'If PD explicitly requested a custom image name, we can safely assume the calling routine is NOT loading a generic image file
    ' from disk - instead, this image came from a scanner, or screen capture, or some format that doesn't automatically yield a
    ' usable filename.
    
    'Therefore, our job is to coordinate between the image's suggested name (which will be suggested at first-save), the actual
    ' location on disk (which we treat as "non-existent", even though we're loading from a temp file of some sort), and the image's
    ' save state (which we forcibly set to FALSE to ensure the user is prompted to save before closing the image).
    Dim cFile As pdFSO
    Set cFile = New pdFSO
            
    If Len(suggestedFilename) = 0 Then
    
        'The calling routine didn't specify a custom image name, so we can assume this is a normal image file.
        'Prep all default attributes using the filename itself.
        targetImage.imgStorage.AddEntry "CurrentLocationOnDisk", srcFile
        targetImage.imgStorage.AddEntry "OriginalFileName", cFile.GetFilename(srcFile, True)
        targetImage.imgStorage.AddEntry "OriginalFileExtension", cFile.GetFileExtension(srcFile)
        
        'Note the image's save state; PDI files are specially marked as having been "saved losslessly".
        If targetImage.currentFileFormat = PDIF_PDI Then
            targetImage.SetSaveState True, pdSE_SavePDI
        Else
            targetImage.SetSaveState True, pdSE_SaveFlat
        End If
        
    Else
    
        'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
        ' dialog in the future by not specifying a location on disk
        targetImage.imgStorage.AddEntry "CurrentLocationOnDisk", ""
        targetImage.imgStorage.AddEntry "OriginalFileName", suggestedFilename
        targetImage.imgStorage.AddEntry "OriginalFileExtension", ""
        
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
    g_AllowViewportRendering = False
    FormMain.mainCanvas(0).SetScrollValue PD_BOTH, 0
    
    'Reflow any image-window-specific display elements on the actual image form (status bar, rulers, etc).
    ' Importantly, this also shows/hides the image tabstrip that's available when multiple images are loaded.
    FormMain.mainCanvas(0).AlignCanvasView
    
    'If the user wants us to resize the image to fit on-screen, do that now
    If (g_AutozoomLargeImages = 0) Then
        
        'Normally we would want to re-enable "g_AllowViewportRendering", but the FitImageToViewport function handles that
        ' reinitialization internally.
        FitImageToViewport
    
    'If the "view 100%" option is checked instead, reset the zoom listbox to match and paint the main window immediately
    Else
        FormMain.mainCanvas(0).SetZoomDropDownIndex srcImage.currentZoomValue
        g_AllowViewportRendering = True
        Viewport_Engine.Stage1_InitializeBuffer srcImage, FormMain.mainCanvas(0), VSR_ResetToZero
    End If
    
    'Notify the UI manager that it now has one more image to deal with
    If (MacroStatus <> MacroBATCH) Then Interface.NotifyImageAdded srcImage.imageID
                            
    'Add this file to the MRU list (unless specifically told not to)
    If addToRecentFiles And (MacroStatus <> MacroBATCH) Then g_RecentFiles.MRU_AddNewFile srcFile, srcImage
    
End Sub

'Legacy import functions for old PDI versions are found below.  These functions are no longer maintained; use at your own risk.
Private Function LoadPDI_Legacy(ByVal pdiPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Legacy PDI file identified.  Starting pdPackage decompression..."
    #End If
    
    On Error GoTo LoadPDIFail
    
    'PDI files require a parent pdImage container
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    pdiReader.Init_ZLib "", True, g_ZLibEnabled
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.ReadPackageFromFile(pdiPath, PD_IMAGE_IDENTIFIER) Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "pdPackage successfully read and initialized.  Starting package parsing..."
        #End If
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.GetNodeDataByIndex(0, True, retBytes, sourceIsUndoFile) Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Initial PDI node retrieved.  Initializing corresponding pdImage object..."
            #End If
            
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
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "pdImage created successfully.  Moving on to individual layers..."
        #End If
        
        'With the main pdImage now assembled, the next task is to populate all layers with two pieces of information:
        ' 1) The layer header, which contains stuff like layer name, opacity, blend mode, etc
        ' 2) Layer-specific information, which varies by layer type.  For DIBs, this will be a raw stream of bytes
        '    containing the layer DIB's raster data.  For text or other vector layers, this is an XML stream containing
        '    whatever information is necessary to construct the layer from scratch.
        
        Dim i As Long
        For i = 0 To dstImage.GetNumOfLayers - 1
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Retrieving layer header " & i & "..."
            #End If
        
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
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Raster layer identified.  Retrieving pixel bits..."
                #End If
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstImage.GetLayerByIndex(i).layerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                'At present, all pdPackage layers will contain premultiplied alpha, so force the corresponding state now
                dstImage.GetLayerByIndex(i).layerDIB.SetInitialAlphaPremultiplicationState True
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByIndex_UnsafeDstPointer(i + 1, False, tmpDIBPointer, sourceIsUndoFile)
            
            'Text and other vector layers
            ElseIf dstImage.GetLayerByIndex(i).IsLayerVector Then
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Vector layer identified.  Retrieving layer XML..."
                #End If
                
                If pdiReader.GetNodeDataByIndex(i + 1, False, retBytes, sourceIsUndoFile) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager gained full Unicode compatibility.
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
                dstImage.NotifyImageChanged UNDO_LAYER, i
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        Next i
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "All layers loaded.  Looking for remaining non-essential PDI data..."
        #End If
        
        'Finally, check to see if the PDI image has a metadata entry.  If it does, load that data now.
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", True, retBytes, sourceIsUndoFile) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Raw metadata chunk found.  Retrieving now..."
            #End If
        
            'Copy the received bytes into a string
            If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            If Not dstImage.imgMetadata.LoadAllMetadata(retString, dstImage.imageID) Then
                
                'For invalid metadata, do not reject the rest of the PDI file.  Instead, just warn the user and carry on.
                Debug.Print "PDI Metadata Node rejected by metadata parser."
                
            End If
        
        End If
        
        '(As of v7.0, a serialized copy of the image's metadata is also stored.  This copy contains all user edits
        ' and other changes.)
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", False, retBytes, sourceIsUndoFile) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Serialized metadata chunk found.  Retrieving now..."
            #End If
        
            'Copy the received bytes into a string
            If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            dstImage.imgMetadata.RecreateFromSerializedXMLData retString
        
        End If
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "PDI parsing complete.  Returning control to main image loader..."
        #End If
        
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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    #End If
    
    'Before falling back to a generic error message, check for a couple known problem states.
    
    'Case 1: zLib is required for this file, but the user doesn't have the zLib plugin
    If pdiReader.GetPackageFlag(PDP_FLAG_ZLIB_REQUIRED, PDP_LOCATION_ANY) And (Not g_ZLibEnabled) Then
        PDMsgBox "The PDI file ""%1"" contains compressed data, but the zLib plugin is missing or disabled." & vbCrLf & vbCrLf & "To enable support for compressed PDI files, click Help > Check for Updates, and when prompted, allow PhotoDemon to download all recommended plugins.", vbInformation + vbOKOnly + vbApplicationModal, "zLib plugin missing", GetFilename(pdiPath)
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
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    pdiReader.Init_ZLib "", True, g_ZLibEnabled
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
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
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    pdiReader.Init_ZLib "", True, g_ZLibEnabled
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
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
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager was given Unicode compatibility.
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
