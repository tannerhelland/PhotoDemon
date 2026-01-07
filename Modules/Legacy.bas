Attribute VB_Name = "Legacy"
'***************************************************************************
'Legacy PhotoDemon support functions
'Copyright 2020-2026 by Tanner Helland
'Created: 25/February/20
'Last updated: 25/February/20
'Last update: migrate legacy functions from elsewhere to this dedicated legacy module
'
'PhotoDemon has existed for multiple decades!
'
'During that time, many (if not *all*) decisions have been revisited multiple times.
' Sometimes, major program features have even been scrapped or rewritten from scratch.
'
'Despite this, I try hard to ensure that any files created by PD are still useable even
' when those formats are reworked.  This module exists to collect legacy file format support
' functions.
'
'IMPORTANTLY: nothing in this module should be modified except to fix security issues or
' enable basic functionality.  Many of these functions are ugly and poorly constructed,
' and that's okay - I replaced them for a reason!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit


'Legacy import functions for old PDI versions are found below.  These functions are no longer maintained
' or updated, but they are guaranteed to work in the current PD build.  DO NOT modify them except to fix
' security issues!
Private Function LoadPDI_Legacy(ByVal pdiPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean

    PDDebug.LogAction "Legacy PDI file identified.  Starting pdPackage decompression..."
    
    On Error GoTo LoadPDIFail
    
    'PDI files require a parent pdImage container
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'First things first: create a legacy pdPackage instance.  It will handle all the messy business of extracting individual
    ' data bits from the source file.
    Dim pdiReader As pdPackageLegacyV1
    Set pdiReader = New pdPackageLegacyV1
    pdiReader.Init_ZLib
    
    'Load the file.  The class will cache the file contents, so we only have to do this once.
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
                CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.SetHeaderFromXML retString, sourceIsUndoFile
        
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
                    CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), UBound(retBytes) + 1
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
                dstImage.GetLayerByIndex(i).GetLayerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                'At present, all pdPackage layers will contain premultiplied alpha, so force the corresponding state now
                dstImage.GetLayerByIndex(i).GetLayerDIB.SetInitialAlphaPremultiplicationState True
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByIndex_UnsafeDstPointer(i + 1, False, tmpDIBPointer, sourceIsUndoFile)
            
            'Text and other vector layers
            ElseIf dstImage.GetLayerByIndex(i).IsLayerVector Then
                
                PDDebug.LogAction "Vector layer identified.  Retrieving layer XML..."
                
                If pdiReader.GetNodeDataByIndex(i + 1, False, retBytes, sourceIsUndoFile) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after the class gained full Unicode compatibility.
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), UBound(retBytes) + 1
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstImage.GetLayerByIndex(i).SetVectorDataFromXML(retString) Then
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
                CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), UBound(retBytes) + 1
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
        
        '(As of v7.0, a serialized copy of the image's metadata is also stored.
        ' This copy contains all user edits and other changes.)
        If pdiReader.GetNodeDataByName("pdMetadata_Raw", False, retBytes, sourceIsUndoFile) Then
        
            PDDebug.LogAction "Serialized metadata chunk found.  Retrieving now..."
            
            'Copy the received bytes into a string
            If pdiReader.GetPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), UBound(retBytes) + 1
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
        PDDebug.LogAction "Selected file is not in PDI format.  Load abandoned."
        LoadPDI_Legacy = False
    
    End If
    
    Exit Function
    
LoadPDIFail:
    
    PDDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    
    'Before falling back to a generic error message, check for a couple known problem states.
    
    'Case 1: compression is enabled for this file, but the user doesn't have the right compression plugin
    Dim cmpMissing As Boolean
    cmpMissing = pdiReader.GetPackageFlag(PDP_HF2_ZlibRequired, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_libdeflate))
    
    If cmpMissing Then
        PDDebug.LogAction "The PDI file " & Files.FileGetName(pdiPath) & " contains compressed data, but the required plugin is missing or disabled."
        Exit Function
    End If
    
    If (Err.Number = PDP_GENERIC_ERROR) Then
        PDDebug.LogAction "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
    Else
        PDDebug.LogAction "An error occurred during PDI loading: " & Err.Number & " - " & Err.Description
    End If
    
    LoadPDI_Legacy = False
    Exit Function

End Function

'PDI version used until Jan 2020 (v8.0 it was replaced).
Public Function LoadPDI_LegacyV2(ByVal pdiPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean
    
    PDDebug.LogAction "PDI file identified.  Starting pdPackage decompression..."
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    On Error GoTo LoadPDIFail
    
    'PDI files require a parent pdImage container
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackageLegacyV2
    Set pdiReader = New pdPackageLegacyV2
    
    'Load the file.  Note that this step will also validate the incoming file.
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
            CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), retSize
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.SetHeaderFromXML retString, sourceIsUndoFile
        
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
                CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), retSize
                
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
                dstImage.GetLayerByIndex(i).GetLayerDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                'At present, all pdPackage layers will contain premultiplied alpha, so force the corresponding state now
                dstImage.GetLayerByIndex(i).GetLayerDIB.SetInitialAlphaPremultiplicationState True
                
                nodeLoadedSuccessfully = pdiReader.GetNodeDataByIndex_UnsafeDstPointer(i + 1, False, tmpDIBPointer)
            
            'Text and other vector layers
            ElseIf dstImage.GetLayerByIndex(i).IsLayerVector Then
                
                PDDebug.LogAction "Vector layer identified.  Retrieving layer XML..."
                
                If pdiReader.GetNodeDataByIndex(i + 1, False, retBytes, False, retSize) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after the class gained full Unicode compatibility.
                    retString = Space$(retSize \ 2)
                    CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), retSize
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstImage.GetLayerByIndex(i).SetVectorDataFromXML(retString) Then
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
            CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), retSize
            
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
            CopyMemoryStrict StrPtr(retString), VarPtr(retBytes(0)), retSize
            
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
        LoadPDI_LegacyV2 = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed.  This may be a legacy PDI file -- try that function next.
        PDDebug.LogAction "Legacy PDI file encountered; dropping back to pdPackage v1 functions..."
        LoadPDI_LegacyV2 = LoadPDI_Legacy(pdiPath, dstDIB, dstImage, sourceIsUndoFile)
    
    End If
    
    Exit Function
    
LoadPDIFail:
    
    PDDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    
    'Before falling back to a generic error message, check for a couple known problem states.
    
    'Case 1: this file is compressed (using one or more libraries), and the user has somehow messed up their PD plugin situation
    Dim cmpMissing As Boolean
    cmpMissing = pdiReader.GetPackageFlag(PDP_HF2_ZlibRequired, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_libdeflate))
    cmpMissing = cmpMissing Or pdiReader.GetPackageFlag(PDP_HF2_ZstdRequired, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_zstd))
    cmpMissing = cmpMissing Or pdiReader.GetPackageFlag(PDP_HF2_Lz4Required, PDP_LOCATION_ANY) And (Not PluginManager.IsPluginCurrentlyEnabled(CCP_lz4))
    
    If cmpMissing Then
        PDDebug.LogAction "The PDI file " & Files.FileGetName(pdiPath) & " contains compressed data, but the required plugin is missing or disabled."
        Exit Function
    End If
    
    If (Err.Number = PDP_GENERIC_ERROR) Then
        PDDebug.LogAction "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
    Else
        PDDebug.LogAction "An error occurred during PDI loading: " & Err.Number & " - " & Err.Description
    End If
    
    LoadPDI_LegacyV2 = False
    Exit Function

End Function

