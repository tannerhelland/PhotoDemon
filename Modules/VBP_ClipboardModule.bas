Attribute VB_Name = "Clipboard_Handler"
'***************************************************************************
'Clipboard Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 15/April/01
'Last updated: 06/September/12
'Last update: rewrote copy/paste against the new layer class.
'
'Module for handling all Windows clipboard routines.  Copy and Paste are the real stars; Cut is not included
' (as there is no purpose for it at present), though Empty Clipboard does make an appearance.
'
'***************************************************************************

Option Explicit

'API functions used to extract file names from clipboard data
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long

Private Const CF_HDROP As Long = 15
Private Const CLIPBOARD_FORMAT_BMP As Long = 2
Private Const CLIPBOARD_FORMAT_METAFILE As Long = 3

'Copy image
Public Sub ClipboardCopy()
    
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    
    'Check for an active selection
    If pdImages(CurrentImage).selectionActive Then
    
        'Fill the temporary layer with the selection
        tmpLayer.createBlank pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
        BitBlt tmpLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.selLeft, pdImages(CurrentImage).mainSelection.selTop, vbSrcCopy
    
        'If the selection contains transparency, blend it against a white background
        If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.compositeBackgroundColor 255, 255, 255
        
    Else
    
        'If a selection is NOT active, just make a copy of the full image
        tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
    
        'If the image contains transparency, blend it against a white background
        If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.compositeBackgroundColor 255, 255, 255
        
    End If
    
    'Copy the temporary layer to the clipboard, then erase it
    tmpLayer.copyLayerToClipboard
    
    tmpLayer.eraseLayer
    
End Sub

'Empty the clipboard
Public Sub ClipboardEmpty()
    Clipboard.Clear
End Sub

'Paste an image (e.g. create a new image based on data in the clipboard
Public Sub ClipboardPaste()
    
    'Make sure the clipboard format is a bitmap
    If Clipboard.GetFormat(CLIPBOARD_FORMAT_BMP) Then
        
        'Copy the image into an StdPicture object
        Dim tmpPicture As StdPicture
        Set tmpPicture = Clipboard.GetData(2)
        
        'Create a temporary layer and copy the temporary StdPicture object into it
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        tmpLayer.CreateFromPicture tmpPicture
        
        'Ask the layer to write its contents to file in BMP format
        Dim tmpClipboardFile As String
        tmpClipboardFile = g_UserPreferences.getTempPath & "PDClipboard.tmp"
        tmpLayer.writeToBitmapFile tmpClipboardFile
        
        'Now that the image is saved on the hard drive, we can delete our temporary objects
        Set tmpPicture = Nothing
        tmpLayer.eraseLayer
        Set tmpLayer = Nothing
        
        'Use the standard image load routine to import the temporary file
        Dim sFile(0) As String
        sFile(0) = tmpClipboardFile
            
        PreLoadImage sFile, False, "Clipboard Image", "Clipboard Image (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
            
        'Be polite and remove the temporary file
        If FileExist(tmpClipboardFile) Then Kill tmpClipboardFile
            
        Message "Clipboard data imported successfully "
    
    'Next, see if the clipboard contains text.  If it does, it may be a hyperlink - if so, try and load it
    ElseIf Clipboard.GetFormat(vbCFText) And ((Left$(Trim(CStr(Clipboard.GetText)), 7) = "http://") Or (Left$(Trim(CStr(Clipboard.GetText)), 6) = "ftp://")) Then
        
        Message "URL found on clipboard.  Attempting to download image at that location..."
        Dim downloadSuccess As Boolean
        downloadSuccess = FormInternetImport.ImportImageFromInternet(Trim(CStr(Clipboard.GetText)))
        
        'If the download failed, let the user know that hey, at least we tried.
        If downloadSuccess = False Then Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
    
    'Next, see if the clipboard contains one or more files.  If it does, try to load them.
    ElseIf Clipboard.GetFormat(vbCFFiles) Then
    
        Dim listFiles() As String
        listFiles = ClipboardGetFiles()
        
        PreLoadImage listFiles
    
    Else
        MsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbExclamation + vbOKOnly + vbApplicationModal, "Windows Clipboard Error"
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

