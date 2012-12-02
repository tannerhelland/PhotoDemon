Attribute VB_Name = "Custom_Dialog_Handler"
Option Explicit

'Present a dialog box to confirm the closing of an unsaved image
Public Function confirmClose(ByVal formID As Long) As VbMsgBoxResult

    Load dialog_UnsavedChanges
    
    dialog_UnsavedChanges.formID = formID
    dialog_UnsavedChanges.ShowDialog
    
    confirmClose = dialog_UnsavedChanges.DialogResult
    
    Unload dialog_UnsavedChanges
    
    Set dialog_UnsavedChanges = Nothing

End Function

'Present a dialog box, asking the user how they want to deal with a multipage image.
Public Function promptMultiImage(ByVal srcFilename As String, ByVal numOfPages As Long) As VbMsgBoxResult

    Load dialog_MultiImage
    
    dialog_MultiImage.ShowDialog srcFilename, numOfPages
    
    promptMultiImage = dialog_MultiImage.DialogResult
    
    Unload dialog_MultiImage
    
    Set dialog_MultiImage = Nothing

End Function
