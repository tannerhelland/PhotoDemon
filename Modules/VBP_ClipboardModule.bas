Attribute VB_Name = "Clipboard_Handler"
'***************************************************************************
'Clipboard Interface
'Copyright 2001-2015 by Tanner Helland
'Created: 15/April/01
'Last updated: 14/June/14
'Last update: add "paste as new layer" actions to Undo stack
'
'Module for handling all Windows clipboard routines.  Copy and Paste are the real stars; Cut is not included
' (as there is no purpose for it at present), though Empty Clipboard does make an appearance.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API functions used to extract file names from clipboard data
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long

Private Const CF_HDROP As Long = 15

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
    
    'In the future, I'd like to move all file interactions in this function to pdFSO, but for now, only a few actions are covered.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    Dim pasteWasSuccessful As Boolean
    pasteWasSuccessful = False
    
    Dim tmpClipboardFile As String, tmpDownloadFile As String
    Dim sFile(0) As String
    Dim sTitle As String, sFilename As String
        
    'PNGs on the clipboard get preferential treatment, as they preserve transparency data - so check for them first
    Dim clpObject As cCustomClipboard
    Set clpObject = New cCustomClipboard
    
    'See if clipboard data is available in PNG format.  If it is, attempt to load it.
    ' (If successful, this IF block will manually exit the sub upon completion.)
    If clpObject.IsDataAvailableForFormatName(FormMain.hWnd, "PNG") Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "PNG format found on clipboard.  Attempting to retrieve data..."
        #End If
        
        Dim PNGID As Long
        PNGID = clpObject.FormatIDForName(FormMain.hWnd, "PNG")
        
        If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
            Dim PNGData() As Byte
        
            If clpObject.GetBinaryData(PNGID, PNGData) Then
                
                'Dump the PNG data out to file
                tmpClipboardFile = g_UserPreferences.GetTempPath & "PDClipboard.png"
                
                If cFile.SaveByteArrayToFile(PNGData, tmpClipboardFile) Then
                
                    'We can now use the standard image load routine to import the temporary file
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
                    
                    clpObject.ClipboardClose
                    
                    pasteWasSuccessful = True
                    
                Else
                    pasteWasSuccessful = False
                End If
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Could not retrieve PNG binary data.  PNG paste action abandoned."
                #End If
            End If
        
            clpObject.ClipboardClose
        
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Could not open clipboard.  PNG paste action abandoned."
            #End If
        End If
        
    End If
    
    'If no PNG data was found, look for an HTML fragment.  Chrome and Firefox will include an HTML fragment link to any
    ' copied image from within the browser, which we can use to download the image in question.
    ' (If successful, this IF block will manually exit the sub upon completion.)
    If clpObject.IsDataAvailableForFormatName(FormMain.hWnd, "HTML Format") And (Not pasteWasSuccessful) Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "HTML format found on clipboard.  Attempting to retrieve data..."
        #End If
        
        Dim HtmlID As Long
        HtmlID = clpObject.FormatIDForName(FormMain.hWnd, "HTML Format")
        
        If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
            Dim htmlString As String
            If clpObject.GetTextData(HtmlID, htmlString) Then
                
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
                        
                        tmpDownloadFile = FormInternetImport.downloadURLToTempFile(Mid$(htmlString, urlStart, urlEnd - urlStart))
                        
                        'If the download was successful, we can now use the standard image load routine to import the temporary file
                        If Len(tmpDownloadFile) <> 0 Then
                        
                            sFile(0) = tmpDownloadFile
                                    
                            Dim tmpFilename As String
                            tmpFilename = tmpDownloadFile
                            StripFilename tmpFilename
                            
                            'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                            If srcIsMeantAsLayer Then
                                Layer_Handler.loadImageAsNewLayer False, sFile(0), True
                            Else
                                LoadFileAsNewImage sFile, False, tmpFilename, tmpFilename, , , , , , True
                            End If
                            
                            'Delete the temporary file
                            cFile.KillFile tmpDownloadFile
                            
                            Message "Clipboard data imported successfully "
                            
                            clpObject.ClipboardClose
                            
                            'Check for load failure.  If the most recent pdImages object is inactive, it's a safe assumption that
                            ' the load operation failed.  (This isn't foolproof, especially if the user loads a ton of images,
                            ' and subsequently unloads images in an arbitrary order - but given the rarity of this situation, I'm content
                            ' to use this technique for predicting failure.)
                            If Not pdImages(UBound(pdImages)) Is Nothing Then
                                If pdImages(UBound(pdImages)).IsActive Then
                                    pasteWasSuccessful = True
                                    Exit Sub
                                Else
                                    pasteWasSuccessful = False
                                End If
                            Else
                                pasteWasSuccessful = False
                            End If
                        
                        Else
                        
                            'If the download failed, let the user know that hey, at least we tried.
                            Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
                            
                        End If
                    
                    End If
                
                End If
                
        
            End If
        
            clpObject.ClipboardClose
        
        End If
    
    End If
    
    
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
        
    'Next, see if the clipboard contains one or more files.  If it does, try to load them.
    ElseIf Clipboard.GetFormat(vbCFFiles) Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "One or more file locations found on clipboard.  Attempting to retrieve data..."
        #End If
        
        Dim listFiles() As String
        listFiles = ClipboardGetFiles()
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If srcIsMeantAsLayer Then
            Dim i As Long
            
            For i = 0 To UBound(listFiles)
                Layer_Handler.loadImageAsNewLayer False, listFiles(i), , True
            Next i
            
        Else
            LoadFileAsNewImage listFiles
        End If
        
        pasteWasSuccessful = True
        
    End If
    
    'If a paste operation was successful, switch the current tool to the layer move/resize tool, which is most likely needed after a
    ' new layer has been pasted.
    If pasteWasSuccessful Then
        If srcIsMeantAsLayer Then toolbar_Toolbox.selectNewTool NAV_MOVE
    Else
        pdMsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbExclamation + vbOKOnly + vbApplicationModal, "Windows Clipboard Error"
    End If
    
End Sub

'The code in the function below is a heavily modified version of code originally located at:
' http://www.vb-helper.com/howto_track_clipboard.html (link still good as of 21 December '12)
' Many thanks to the original author(s).
Public Function ClipboardGetFiles() As String()
    
    Dim drop_handle As Long
    Dim num_file_names As Long
    Dim file_names() As String
    Dim file_name As String * 1024
    Dim i As Long

    ' Make sure there is file data.
    If Clipboard.GetFormat(vbCFFiles) Then
        
        ' File data exists. Get it.
        ' Open the clipboard.
        If OpenClipboard(0) Then
            ' The clipboard is open.

            ' Get the handle to the dropped list of files.
            drop_handle = GetClipboardData(CF_HDROP)

            ' Get the number of dropped files.
            num_file_names = DragQueryFile(drop_handle, -1, vbNullString, 0)

            ' Get the file names.
            ReDim file_names(0 To num_file_names - 1) As String
            For i = 0 To num_file_names - 1
                ' Get the file name.
                DragQueryFile drop_handle, i, file_name, Len(file_name)

                ' Truncate at the NULL character.
                file_names(i) = Left$(file_name, InStr(file_name, vbNullChar) - 1)
            Next i

            ' Close the clipboard.
            CloseClipboard

            ' Assign the return value.
            ClipboardGetFiles = file_names
            
        End If
        
    End If
    
End Function

'Because OLE drag/drop is so similar to clipboard functionality, I have included that functionality here.
' Data and Effect are passed as-is from the calling function, while intendedTargetIsLayer controls whether the image file(s)
' (if valid) should be loaded as new layers or new images.
Public Function loadImageFromDragDrop(ByRef Data As DataObject, ByRef Effect As Long, ByVal intendedTargetIsLayer As Boolean) As Boolean

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
            
            loadImageFromDragDrop = True
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
        
        loadImageFromDragDrop = True
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
                loadImageFromDragDrop = True
                Exit Function
            
            Else
            
                'If the download failed, let the user know that hey, at least we tried.
                Message "Image download failed.  Please supply a valid image URL and try again."
                
            End If
            
        End If
    
    End If
    
    'If we made it all the way here, something went horribly wrong
    loadImageFromDragDrop = False
    Exit Function

End Function
