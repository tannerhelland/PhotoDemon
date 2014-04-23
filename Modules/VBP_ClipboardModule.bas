Attribute VB_Name = "Clipboard_Handler"
'***************************************************************************
'Clipboard Interface
'Copyright ©2001-2014 by Tanner Helland
'Created: 15/April/01
'Last updated: 19/August/13
'Last update: removed all references to metafile handling; there are no plans for me to restore EMF/WMF clipboard support
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
Private Const CLIPBOARD_FORMAT_BMP As Long = 2

'Copy the current layer (or composite image, if copyMerged is true) to the clipboard.
' If a selection is active, crop the image to the layer area first.
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
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.fixPremultipliedAlpha False
            
        End If
        
    End If
    
    'Copy the temporary DIB to the clipboard, then erase it
    tmpDIB.copyDIBToClipboard
    tmpDIB.eraseDIB
    
End Sub

'Empty the clipboard
Public Sub ClipboardEmpty()
    Clipboard.Clear
End Sub

'Paste an image (e.g. create new image data based on whatever is in the clipboard).
' The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image, or as a new layer in an existing image.
Public Sub ClipboardPaste(ByVal srcIsMeantAsLayer As Boolean)
    
    Dim tmpClipboardFile As String
    Dim sFile(0) As String
    Dim sTitle As String, sFilename As String
    
    'PNGs on the clipboard get preferential treatment, as they preserve transparency data - so check for them first
    Dim clpObject As cCustomClipboard
    Set clpObject = New cCustomClipboard
    
    'See if clipboard data is available in PNG format.  If it is, attempt to load it
    If clpObject.IsDataAvailableForFormatName(FormMain.hWnd, "PNG") Then
            
        Dim PNGID As Long
        PNGID = clpObject.FormatIDForName(FormMain.hWnd, "PNG")
        
        Dim PNGData() As Byte
        If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
            If clpObject.GetBinaryData(PNGID, PNGData) Then
                
                'Dump the PNG data out to file
                tmpClipboardFile = g_UserPreferences.getTempPath & "PDClipboard.png"
                
                Dim fileID As Integer
                fileID = FreeFile()
                Open tmpClipboardFile For Binary As #fileID
                    Put #fileID, 1, PNGData
                Close #fileID
                
                'We can now use the standard image load routine to import the temporary file
                sFile(0) = tmpClipboardFile
                sTitle = g_Language.TranslateMessage("Clipboard Image")
                sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If srcIsMeantAsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle
                Else
                    LoadFileAsNewImage sFile, False, sTitle, sFilename
                End If
                    
                'Be polite and remove the temporary file
                If FileExist(tmpClipboardFile) Then Kill tmpClipboardFile
                    
                Message "Clipboard data imported successfully "
                
                clpObject.ClipboardClose
                Exit Sub
        
            End If
        
            clpObject.ClipboardClose
        
        End If
                
    End If
    
    'Make sure the clipboard format is a bitmap
    If Clipboard.GetFormat(CLIPBOARD_FORMAT_BMP) Then
        
        'Copy the image into an StdPicture object
        Dim tmpPicture As StdPicture
        Set tmpPicture = Clipboard.GetData(2)
        
        'Create a temporary DIB and copy the temporary StdPicture object into it
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromPicture tmpPicture
        
        'Ask the DIB to write its contents to file in BMP format
        tmpClipboardFile = g_UserPreferences.getTempPath & "PDClipboard.tmp"
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
            Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle
        Else
            LoadFileAsNewImage sFile, False, sTitle, sFilename
        End If
            
        'Be polite and remove the temporary file
        If FileExist(tmpClipboardFile) Then Kill tmpClipboardFile
            
        Message "Clipboard data imported successfully "
    
    'Next, see if the clipboard contains text.  If it does, it may be a hyperlink - if so, try and load it.
    ' TODO: make hyperlinks work with "Paste as new layer".  Right now they will always default to loading as a new image.
    ElseIf Clipboard.GetFormat(vbCFText) And ((Left$(Trim(Str(Clipboard.GetText)), 7) = "http://") Or (Left$(Trim(Str(Clipboard.GetText)), 6) = "ftp://")) Then
        
        Message "URL found on clipboard.  Attempting to download image at that location..."
        Dim downloadSuccess As Boolean
        downloadSuccess = FormInternetImport.ImportImageFromInternet(Trim(Str(Clipboard.GetText)))
        
        'If the download failed, let the user know that hey, at least we tried.
        If downloadSuccess = False Then Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
    
    'Next, see if the clipboard contains one or more files.  If it does, try to load them.
    ElseIf Clipboard.GetFormat(vbCFFiles) Then
    
        Dim listFiles() As String
        listFiles = ClipboardGetFiles()
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If srcIsMeantAsLayer Then
            Dim i As Long
            
            For i = 0 To UBound(listFiles)
                Layer_Handler.loadImageAsNewLayer False, listFiles(i)
            Next i
            
        Else
            LoadFileAsNewImage listFiles
        End If
        
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

