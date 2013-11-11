Attribute VB_Name = "Custom_Dialog_Handler"
'***************************************************************************
'Custom Dialog Interface
'Copyright ©2011-2013 by Tanner Helland
'Created: 30/November/12
'Last updated: 12/February/13
'Last update: added support for the new language selection dialog
'
'Module for handling all custom dialog forms used by PhotoDemon.  There are quite a few already, and I expect
' the number to grow as I phase out generic message boxes in favor of more descriptive (and usable) dialogs
' designed around a specific purpose.
'
'All dialogs are based off the same template, as you can see - they are just modal forms with a specially
' designed ".ShowDialog" sub or function that sets a ".userResponse" property.  The wrapper function in this
' module simply checks that value, unloads the dialog form, then returns the value; this keeps all load/unload
' burdens here so that calling functions can simply use a MsgBox-style line to call the dialogs and check
' the user's response.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Display a custom color selection dialog
Public Function choosePDColor(ByVal oldColor As Long, ByRef newColor As Long) As VbMsgBoxResult

    Load dialog_ColorSelector
    dialog_ColorSelector.showDialog oldColor
    
    choosePDColor = dialog_ColorSelector.DialogResult
    If choosePDColor = vbOK Then newColor = dialog_ColorSelector.newColor
    
    Unload dialog_ColorSelector
    Set dialog_ColorSelector = Nothing

End Function

'Present a dialog box to confirm the closing of an unsaved image
Public Function confirmClose(ByVal formID As Long) As VbMsgBoxResult

    Load dialog_UnsavedChanges
    
    dialog_UnsavedChanges.formID = formID
    dialog_UnsavedChanges.showDialog pdImages(formID).containingForm
    
    confirmClose = dialog_UnsavedChanges.DialogResult
    
    Unload dialog_UnsavedChanges
    Set dialog_UnsavedChanges = Nothing

End Function

'Present a dialog box to ask the user how they want to deal with a multipage image.
Public Function promptMultiImage(ByVal srcFilename As String, ByVal numOfPages As Long) As VbMsgBoxResult

    Load dialog_MultiImage
    dialog_MultiImage.showDialog srcFilename, numOfPages
    
    promptMultiImage = dialog_MultiImage.DialogResult
    
    Unload dialog_MultiImage
    Set dialog_MultiImage = Nothing

End Function

'Present a dialog box to ask the user for various JPEG export settings
Public Function promptJPEGSettings(ByRef srcImage As pdImage) As VbMsgBoxResult

    Load dialog_ExportJPEG
    Set dialog_ExportJPEG.imageBeingExported = srcImage
    dialog_ExportJPEG.showDialog

    promptJPEGSettings = dialog_ExportJPEG.DialogResult
    
    Set dialog_ExportJPEG.imageBeingExported = Nothing
    
    Unload dialog_ExportJPEG
    Set dialog_ExportJPEG = Nothing

End Function

'Present a dialog box to ask the user for various JPEG-2000 (JP2) export settings
Public Function promptJP2Settings() As VbMsgBoxResult

    Load dialog_ExportJP2
    dialog_ExportJP2.showDialog

    promptJP2Settings = dialog_ExportJP2.DialogResult
    
    Unload dialog_ExportJP2
    Set dialog_ExportJP2 = Nothing

End Function

'Present a dialog box to ask the user for desired output color depth
Public Function promptColorDepth(ByVal outputFormat As Long) As VbMsgBoxResult

    Load dialog_ExportColorDepth
    dialog_ExportColorDepth.imageFormat = outputFormat
    dialog_ExportColorDepth.showDialog

    promptColorDepth = dialog_ExportColorDepth.DialogResult
    
    Unload dialog_ExportColorDepth
    Set dialog_ExportColorDepth = Nothing

End Function

'Present a dialog box to ask the user for an alpha-cutoff value.  This is used when reducing a complex (32bpp)
' alpha channel to a simple (8bpp) one.
Public Function promptAlphaCutoff(ByRef srcLayer As pdLayer) As VbMsgBoxResult

    Load dialog_AlphaCutoff
    dialog_AlphaCutoff.refLayer = srcLayer
    dialog_AlphaCutoff.showDialog

    promptAlphaCutoff = dialog_AlphaCutoff.DialogResult
    
    Unload dialog_AlphaCutoff
    Set dialog_AlphaCutoff = Nothing

End Function

'If the user is running in the IDE, warn them of the consequences of doing so
Public Function displayIDEWarning() As VbMsgBoxResult

    Load dialog_IDEWarning

    dialog_IDEWarning.showDialog

    displayIDEWarning = dialog_IDEWarning.DialogResult
    
    Unload dialog_IDEWarning
    Set dialog_IDEWarning = Nothing

End Function

