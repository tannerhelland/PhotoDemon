Attribute VB_Name = "Clipboard_Handler"
'***************************************************************************
'Clipboard Interface
'Copyright 2001-2015 by Tanner Helland
'Created: 15/April/01
'Last updated: 10/November/15
'Last update: start overhaul to integrate new pdClipboard class
'
'Module for handling all Windows clipboard routines.  Copy and Paste are the real stars; Cut is not included
' (as there is no purpose for it at present), though Empty Clipboard does make an appearance.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Copy the current selection (or entire layer, if no selection is active) to the clipboard, then erase the selected area
' (or entire layer, if no selection is active).
Public Sub ClipboardCut(ByVal cutMerged As Boolean)

    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Check for an active selection
    If pdImages(g_CurrentImage).selectionActive Then
    
        'Fill the temporary DIB with the selection
        pdImages(g_CurrentImage).retrieveProcessedSelection tmpDIB, False, cutMerged
        
    Else
        
        'If a selection is NOT active, just make a copy of the full layer or image, depending on the merged request
        If cutMerged Then
            pdImages(g_CurrentImage).getCompositedImage tmpDIB, False
        Else
            tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveLayer.layerDIB
            
            'Layers are always premultiplied, so we must unpremultiply it now if 32bpp
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.setAlphaPremultiplication False
            
        End If
        
    End If
    
    'Copy the temporary DIB to the clipboard, then erase it
    DIB_Handler.copyDIBToClipboard tmpDIB
    Set tmpDIB = Nothing
    
    'Now, we have the added step of erasing the selected area from the screen.  "Cut merged" requires us to delete the selected
    ' region from all visible layers, so vary the loop bounds accordingly.
    Dim startLayer As Long, endLayer As Long
    
    If cutMerged Then
        startLayer = 0
        endLayer = pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        startLayer = pdImages(g_CurrentImage).getActiveLayerIndex
        endLayer = pdImages(g_CurrentImage).getActiveLayerIndex
    End If
    
    Dim i As Long
    For i = startLayer To endLayer
        
        'For "cut merged", ignore transparent layers
        If cutMerged Then
        
            If pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
                
                'If a selection is active, erase the selected area.  Otherwise, wipe the whole layer.
                If pdImages(g_CurrentImage).selectionActive Then
                    pdImages(g_CurrentImage).eraseProcessedSelection i
                Else
                    Layer_Handler.eraseLayerByIndex i
                End If
                
            End If
        
        'For "cut from layer", erase the selection regardless of layer visibility
        Else
        
            'If a selection is active, erase the selected area.  Otherwise, delete the given layer.
            If pdImages(g_CurrentImage).selectionActive Then
                pdImages(g_CurrentImage).eraseProcessedSelection i
            Else
                Layer_Handler.deleteLayer i
            End If
            
        End If
        
    Next i
    
    'Redraw the active viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Copy the current layer (or composite image, if copyMerged is true) to the clipboard.
Public Sub ClipboardCopy(ByVal copyMerged As Boolean)
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Check for an active selection
    If pdImages(g_CurrentImage).selectionActive Then
    
        'Fill the temporary DIB with the selection
        pdImages(g_CurrentImage).retrieveProcessedSelection tmpDIB, False, copyMerged
        
    Else
    
        'If a selection is NOT active, just make a copy of the full layer or image, depending on the merged request
        If copyMerged Then
            pdImages(g_CurrentImage).getCompositedImage tmpDIB, False
        Else
            tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveLayer.layerDIB
            
            'Layers are always premultiplied, so we must unpremultiply it now if 32bpp
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.setAlphaPremultiplication False
            
        End If
        
    End If
    
    'Copy the temporary DIB to the clipboard, then erase it
    DIB_Handler.copyDIBToClipboard tmpDIB
    tmpDIB.eraseDIB
    
End Sub

'Empty the clipboard
Public Sub ClipboardEmpty()
    Clipboard.Clear
End Sub

'Paste an image (e.g. create new image data based on whatever is in the clipboard).
' The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image, or as a new layer in an existing image.
Public Sub ClipboardPaste(ByVal srcIsMeantAsLayer As Boolean)
    
    Dim pasteWasSuccessful As Boolean
    pasteWasSuccessful = False
    
    'Prep a bunch of generic clipboard handling variables
    Dim tmpClipboardFile As String, tmpDownloadFile As String
    Dim sFile(0) As String
    Dim sTitle As String, sFilename As String
    
    'Also, note that all file interactions occur through pdFSO, so we can support Unicode filenames/paths.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Attempt to open the clipboard
    Dim clpObject As pdClipboard
    Set clpObject = New pdClipboard
    If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
        'When debugging, it's nice to know what clipboard formats the OS reports prior to actually retrieving them.
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Clipboard reports the following formats: " & clpObject.GetListOfAvailableFormatNames()
        #End If
        
        'PNGs on the clipboard get preferential treatment, as they preserve transparency data - so check for them first.
        If clpObject.DoesClipboardHaveFormatName("PNG") Then
            pasteWasSuccessful = Clipboard_Handler.ClipboardPaste_CustomImageFormat(clpObject, "PNG", srcIsMeantAsLayer, "png")
        End If
        
        'If we couldn't find PNG data (or something went horribly wrong during that step), look for an HTML fragment next.
        ' Images copied from web browsers typically create an HTML fragment, which should have a direct link to the copied image.
        '  Downloading the image manually lets us maintain things like ICC profiles and the image's original filename.
        If clpObject.DoesClipboardHaveHTML() And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = Clipboard_Handler.ClipboardPaste_HTML(clpObject, srcIsMeantAsLayer)
        End If
        
        'JPEGs are another possibility.  We prefer them less than PNG or direct download (because there's no guarantee that the
        ' damn browser didn't re-encode them, but they're better than bitmaps or DIBs because they may retain metadata and
        ' color profiles, so test for JPEG next.  (Also, note that certain versions of Microsoft Office use "JFIF" as the identifier,
        ' for reasons known only to them...)
        If clpObject.DoesClipboardHaveFormatName("JPEG") And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = Clipboard_Handler.ClipboardPaste_CustomImageFormat(clpObject, "JPEG", srcIsMeantAsLayer, "jpg")
        End If
        
        If clpObject.DoesClipboardHaveFormatName("JPG") And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = Clipboard_Handler.ClipboardPaste_CustomImageFormat(clpObject, "JPG", srcIsMeantAsLayer, "jpg")
        End If
        
        If clpObject.DoesClipboardHaveFormatName("JFIF") And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = Clipboard_Handler.ClipboardPaste_CustomImageFormat(clpObject, "JFIF", srcIsMeantAsLayer, "jpg")
        End If
        
         'Next, see if the clipboard contains one or more files.  If it does, try to load them.
        If clpObject.DoesClipboardHaveFiles() And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = Clipboard_Handler.ClipboardPaste_ListOfFiles(clpObject, srcIsMeantAsLayer)
        End If
        
        'Regardless of success or failure, make sure to close the clipboard now that we're done with it.
        clpObject.ClipboardClose
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  Couldn't open the clipboard; is it possible another program has locked it?"
        #End If
    End If
    
    '*** END OF API CLIPBOARD INTERACTIONS ***
    '*** Everything beyond this point uses pure VB code ***
    '*** TODO: FIX THIS! ***
        
    'Make sure the clipboard format is a bitmap
    If Clipboard.GetFormat(vbCFBitmap) And (Not pasteWasSuccessful) Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "BMP format found on clipboard.  Attempting to retrieve data..."
        #End If
        
        'Copy the image into an StdPicture object
        Dim tmpPicture As StdPicture
        Set tmpPicture = Clipboard.GetData(2)
        
        'Create a temporary DIB and copy the temporary StdPicture object into it
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromPicture tmpPicture
        
        'Ask the DIB to write its contents to file in BMP format
        tmpClipboardFile = g_UserPreferences.GetTempPath & "PD_Clipboard.tmp"
        tmpDIB.writeToBitmapFile tmpClipboardFile
        
        'Now that the image is saved on the hard drive, we can delete our temporary objects
        Set tmpPicture = Nothing
        tmpDIB.eraseDIB
        Set tmpDIB = Nothing
        
        'Use the standard image load routine to import the temporary file
        sFile(0) = tmpClipboardFile
        sTitle = g_Language.TranslateMessage("Clipboard Image")
        sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If srcIsMeantAsLayer Then
            Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle, True
        Else
            LoadFileAsNewImage sFile, False, sTitle, sFilename
        End If
            
        'Be polite and remove the temporary file
        cFile.KillFile tmpClipboardFile
            
        Message "Clipboard data imported successfully "
        
        pasteWasSuccessful = True
    
    'Next, see if the clipboard contains text.  If it does, it may be a hyperlink - if so, try and load it.
    ElseIf Clipboard.GetFormat(vbCFText) And (Not pasteWasSuccessful) Then
        
        tmpDownloadFile = Trim$(Clipboard.GetText)
        
        If (StrComp(UCase$(Left$(tmpDownloadFile, 4)), "HTTP", vbBinaryCompare) = 0) Or (StrComp(UCase$(Left$(tmpDownloadFile, 6)), "FTP://", vbBinaryCompare) = 0) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Probably URL found on clipboard.  Attempting to retrieve data..."
            #End If
            
            Message "Image URL found on clipboard.  Attempting to download..."
            
            tmpDownloadFile = FormInternetImport.downloadURLToTempFile(tmpDownloadFile)
            
            'If the download was successful, we can now use the standard image load routine to import the temporary file
            If Len(tmpDownloadFile) <> 0 Then
            
                sFile(0) = tmpDownloadFile
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If srcIsMeantAsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, sFile(0), , True
                Else
                    LoadFileAsNewImage sFile, False, , GetFilename(tmpDownloadFile)
                End If
                
                'Delete the temporary file
                cFile.KillFile tmpDownloadFile
                
                Message "Clipboard data imported successfully "
        
                clpObject.ClipboardClose
                
                pasteWasSuccessful = True
            
            Else
            
                'If the download failed, let the user know that hey, at least we tried.
                Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
                
            End If
            
        End If
        
    End If
        
    'If a paste operation was successful, switch the current tool to the layer move/resize tool, which is most likely needed after a
    ' new layer has been pasted.
    If pasteWasSuccessful Then
        If srcIsMeantAsLayer Then toolbar_Toolbox.selectNewTool NAV_MOVE
    Else
        PDMsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbExclamation + vbOKOnly + vbApplicationModal, "Windows Clipboard Error"
    End If
    
End Sub

'If the clipboard contains custom-format image data (most commonly PNG or JPEG), you can call this function to initiate a "paste" command
' using the custom image data as a source.  The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image,
' or as a new layer in an existing image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_CustomImageFormat(ByRef clpObject As pdClipboard, ByVal clipboardFormatName As String, ByVal srcIsMeantAsLayer As Boolean, Optional ByVal tmpFileExtension As String = "tmp") As Boolean
        
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_CustomImageFormat = False
    
    'All paste operations use a few consistent variables
    
    'Raw retrieval storage variables
    Dim clipFormatID As Long, rawClipboardData() As Byte
    
    'Temporary file for storing the clipboard data.  (This lets us use PD's central image load function.)
    Dim tmpClipboardFile As String
    
    'Additional file information variables, which we pass to the central load function to let it know that this is only a temp file,
    ' and it should use these hint values instead of assuming normal image load behavior.
    Dim sFile() As String, sTitle As String, sFilename As String
    ReDim sFile(0) As String
    
    'pdFSO is used to ensure Unicode subfolder compatibility
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If clpObject.DoesClipboardHaveFormatName(clipboardFormatName) Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ClipboardPaste_CustomImageFormat() will now attempt to load " & clipboardFormatName & " from the clipboard..."
        #End If
        
        'Because custom-format image data can be registered by many programs, retrieve this image format's unique ID now.
        clipFormatID = clpObject.GetFormatIDFromName(clipboardFormatName)
        
        'Pull the data into a local array
        If clpObject.GetClipboardBinaryData(clipFormatID, rawClipboardData) Then
            
            'Dump the data out to file
            tmpClipboardFile = g_UserPreferences.GetTempPath & "PDClipboard." & tmpFileExtension
            If cFile.SaveByteArrayToFile(rawClipboardData, tmpClipboardFile) Then
                
                'We no longer need our local copy of the clipboard data
                Erase rawClipboardData
                
                'We can now use the standard image load routine to import the temporary file.  Because we don't want the
                ' load function to use the temporary file name as the image name, we manually supply a filename to suggest
                ' if the user eventually tries to save the file.
                sFile(0) = tmpClipboardFile
                sTitle = g_Language.TranslateMessage("Clipboard Image")
                sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If srcIsMeantAsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle, True
                Else
                    LoadFileAsNewImage sFile, False, sTitle, sFilename
                End If
                    
                'Be polite and remove the temporary file
                cFile.KillFile tmpClipboardFile
                
                'If we made it all the way here, the load was (probably?) successful
                Message "Clipboard data imported successfully "
                ClipboardPaste_CustomImageFormat = True
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  Clipboard image data (probably PNG) could not be written to a temp file."
                #End If
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  Clipboard.GetBinaryData failed on custom image data (probably PNG).  Special paste action abandoned."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_CustomImageFormat was called, but the requested data doesn't exist on the clipboard."
        #End If
    End If
    
End Function

'If the clipboard contains HTML text (presumably copied from a web browser), you can call this function to initiate a "paste" command
' using the HTML text as a source.  When copying an image from the web, most web browsers will include a link to the original image
' on the clipboard; we prefer to download this vs grabbing the actual image bits, as we provide much more comprehensive handling for
' things like metadata, special PNG chunks, ICC profiles, and more.
'Also, the parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image, or as a new layer in the
' active image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_HTML(ByRef clpObject As pdClipboard, ByVal srcIsMeantAsLayer As Boolean) As Boolean
    
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_HTML = False
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If clpObject.DoesClipboardHaveHTML() Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ClipboardPaste_HTML() will now attempt to find valid image HTML on the clipboard..."
        #End If
        
        'Pull the HTML data into a local string
        Dim htmlString As String
        If clpObject.GetClipboardHTML(htmlString) Then
            
            'Look for an image reference within the HTML snippet
            If InStr(1, UCase$(htmlString), "<IMG ", vbBinaryCompare) > 0 Then
            
                'Retrieve the full image path, which will be between the first set of quotation marks following the
                ' "<img src=" statement in the HTML snippet.
                Dim vbQuoteMark As String
                vbQuoteMark = """"
                
                'Parse out the URL between the img src quotes
                Dim urlStart As Long, urlEnd As Long
                urlStart = InStr(1, UCase$(htmlString), "<IMG ", vbBinaryCompare)
                If urlStart > 0 Then urlStart = InStr(urlStart, UCase$(htmlString), "SRC=", vbBinaryCompare)
                If urlStart > 0 Then urlStart = InStr(urlStart, htmlString, vbQuoteMark, vbBinaryCompare) + 1
                
                'The magic number 6 below is calculated as the length of (src="), + 1 to advance to the
                ' character immediately following the quotation mark.
                If urlStart > 0 Then urlEnd = InStr(urlStart + 6, htmlString, vbQuoteMark, vbBinaryCompare)
                
                'As a failsafe, make sure a valid URL was actually found
                If (urlStart > 0) And (urlEnd > 0) Then
                    
                    Message "Image URL found on clipboard.  Attempting to download..."
                    
                    Dim tmpDownloadFile As String
                    tmpDownloadFile = FormInternetImport.downloadURLToTempFile(Mid$(htmlString, urlStart, urlEnd - urlStart))
                    
                    'pdFSO is used to ensure Unicode filename compatibility
                    Dim cFile As pdFSO
                    Set cFile = New pdFSO
                    
                    'If the download was successful, we can now use the standard image load routine to import the temporary file
                    If cFile.FileLenW(tmpDownloadFile) <> 0 Then
                        
                        'Additional file information variables, which we pass to the central load function to let it know that this is only a temp file,
                        ' and it should use these hint values instead of assuming normal image load behavior.
                        Dim sFile() As String
                        ReDim sFile(0) As String
                        sFile(0) = tmpDownloadFile
                                
                        Dim tmpFilename As String
                        tmpFilename = cFile.GetFilename(tmpDownloadFile, True)
                        
                        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                        If srcIsMeantAsLayer Then
                            Layer_Handler.loadImageAsNewLayer False, sFile(0), True
                        Else
                            LoadFileAsNewImage sFile, False, tmpFilename, tmpFilename, , , , , , True
                        End If
                        
                        'Delete the temporary file
                        cFile.KillFile tmpDownloadFile
                        
                        Message "Clipboard data imported successfully "
                            
                        'Check for load failure.  If the most recent pdImages object is inactive, it's a safe assumption that
                        ' the load operation failed.  (This isn't foolproof, especially if the user loads a ton of images,
                        ' and subsequently unloads images in an arbitrary order - but given the rarity of this situation, I'm content
                        ' to use this technique for predicting failure.)
                        If g_CurrentImage <= UBound(pdImages) Then
                            If Not pdImages(g_CurrentImage) Is Nothing Then
                                If pdImages(g_CurrentImage).IsActive Then
                                    ClipboardPaste_HTML = True
                                End If
                            End If
                        End If
                    
                    Else
                        
                        'If the download failed, let the user know that hey, at least we tried.
                        Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
                        
                    End If
                
                'An image tag was found, but a parsing error occurred when trying to strip out the source URL.  This is okay;
                ' exit immediately without raising any errors.
                Else
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "Clipboard.GetClipboardHTML was successful and an image URL was located, but a parsing error occurred."
                    #End If
                End If
                
            'No image tag found, which is fine; exit immediately without raising any errors.
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Clipboard.GetClipboardHTML was successful, but no image URL found."
                #End If
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  Clipboard.GetClipboardHTML failed.  Special paste action abandoned."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_HTML was called, but HTML data doesn't exist on the clipboard."
        #End If
    End If
    
End Function

'If one or more files exist on the clipboard, attempt to paste them all.
Private Function ClipboardPaste_ListOfFiles(ByRef clpObject As pdClipboard, ByVal srcIsMeantAsLayer As Boolean)
    
    'Make sure files actually exist on the clipboard
    If clpObject.DoesClipboardHaveFiles() Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ClipboardPaste_ListOfFiles() will now attempt to load one or more files from the clipboard..."
        #End If
        
        Dim listOfFiles() As String, numOfFiles As Long
        If clpObject.GetFileList(listOfFiles, numOfFiles) Then
            
            'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
            If srcIsMeantAsLayer Then
                Dim i As Long
                For i = 0 To UBound(listOfFiles)
                    If Len(listOfFiles(i)) <> 0 Then Layer_Handler.loadImageAsNewLayer False, listOfFiles(i), , True
                Next i
            Else
                LoadFileAsNewImage listOfFiles
            End If
            
            ClipboardPaste_ListOfFiles = True
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  ClipboardPaste_ListOfFiles couldn't retrieve a valid file list from pdClipboard."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_ListOfFiles was called, but no file paths exist on the clipboard."
        #End If
    End If
        
End Function

'Because OLE drag/drop is so similar to clipboard functionality, I have included that functionality here.
' Data and Effect are passed as-is from the calling function, while intendedTargetIsLayer controls whether the image file(s)
' (if valid) should be loaded as new layers or new images.
Public Function LoadImageFromDragDrop(ByRef Data As DataObject, ByRef Effect As Long, ByVal intendedTargetIsLayer As Boolean) As Boolean

    Dim sFile() As String
    Dim tmpString As String
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        
        Dim countFiles As Long
        countFiles = 0
                
        For Each oleFilename In Data.Files
            tmpString = oleFilename
            If (Len(tmpString) <> 0) And cFile.FileExist(tmpString) Then
                sFile(countFiles) = tmpString
                countFiles = countFiles + 1
            End If
        Next oleFilename
        
        'Make sure at least one valid, existant file was found
        If countFiles > 0 Then
        
            'Because the OLE drop may include blank strings, verify the size of the array against countFiles
            ReDim Preserve sFile(0 To countFiles - 1) As String
            
            'If an open image exists, pass the list of filenames to LoadImageAsNewLayer, which will load the images one-at-a-time
            Dim i As Long
            
            If (g_OpenImageCount > 0) And intendedTargetIsLayer Then
                For i = 0 To UBound(sFile)
                    Layer_Handler.loadImageAsNewLayer False, sFile(i)
                Next i
            Else
                LoadFileAsNewImage sFile
            End If
            
            LoadImageFromDragDrop = True
            Exit Function
            
        End If
    
    End If
    
    'If the data is not a file list, see if it's a bitmap
    If Data.GetFormat(vbCFBitmap) Then
    
        'Copy the image into an StdPicture object
        Dim tmpPicture As StdPicture
        Set tmpPicture = Data.GetData(vbCFBitmap)
        
        'Create a temporary DIB and copy the temporary StdPicture object into it
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromPicture tmpPicture
        
        'Ask the DIB to write its contents to file in BMP format
        tmpString = g_UserPreferences.GetTempPath & "PD_DragDrop.tmp"
        tmpDIB.writeToBitmapFile tmpString
        
        'Now that the image is saved on the hard drive, we can delete our temporary objects
        Set tmpPicture = Nothing
        tmpDIB.eraseDIB
        Set tmpDIB = Nothing
        
        'Use the standard image load routine to import the temporary file
        Dim sTitle As String, sFilename As String
        ReDim sFile(0) As String
        sFile(0) = tmpString
        sTitle = g_Language.TranslateMessage("Imported Image")
        sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If (g_OpenImageCount > 0) And intendedTargetIsLayer Then
            Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle
        Else
            LoadFileAsNewImage sFile, False, sTitle, sFilename
        End If
            
        'Be polite and remove the temporary file
        If cFile.FileExist(tmpString) Then cFile.KillFile tmpString
            
        Message "Image imported successfully "
        
        LoadImageFromDragDrop = True
        Exit Function
    
    'If the data is not a file list, see if it's a URL.
    ElseIf Data.GetFormat(vbCFText) Then
    
        Dim tmpDownloadFile As String
        tmpDownloadFile = Trim$(Data.GetData(vbCFText))
        
        If (StrComp(UCase$(Left$(tmpDownloadFile, 7)), "HTTP://", vbBinaryCompare) = 0) Or (StrComp(UCase$(Left$(tmpDownloadFile, 8)), "HTTPS://", vbBinaryCompare) = 0) Or (StrComp(UCase$(Left$(tmpDownloadFile, 6)), "FTP://", vbBinaryCompare) = 0) Then
        
            Message "Image URL found on clipboard.  Attempting to download..."
            
            tmpDownloadFile = FormInternetImport.downloadURLToTempFile(tmpDownloadFile)
            
            'If the download was successful, we can now use the standard image load routine to import the temporary file
            If Len(tmpDownloadFile) <> 0 Then
                
                'Depending on the number of open images, load the clipboard data as a new image or as a new layer in the current image
                If (g_OpenImageCount > 0) And intendedTargetIsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, tmpDownloadFile
                Else
                    ReDim sFile(0) As String
                    sFile(0) = tmpDownloadFile
                    LoadFileAsNewImage sFile, False, , GetFilename(tmpDownloadFile)
                End If
                
                'Delete the temporary file
                If cFile.FileExist(tmpDownloadFile) Then cFile.KillFile tmpDownloadFile
                
                'Exit!
                LoadImageFromDragDrop = True
                Exit Function
            
            Else
            
                'If the download failed, let the user know that hey, at least we tried.
                Message "Image download failed.  Please supply a valid image URL and try again."
                
            End If
            
        End If
    
    End If
    
    'If we made it all the way here, something went horribly wrong
    LoadImageFromDragDrop = False
    Exit Function

End Function
