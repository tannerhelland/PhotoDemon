Attribute VB_Name = "Dialog_Handler"
'***************************************************************************
'Custom Dialog Interface
'Copyright ©2012-2014 by Tanner Helland
'Created: 30/November/12
'Last updated: 14/February/14
'Last update: added support for the new WebP export dialog
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
Public Function choosePDColor(ByVal oldColor As Long, ByRef newColor As Long, Optional ByRef callingControl As colorSelector) As VbMsgBoxResult

    Load dialog_ColorSelector
    dialog_ColorSelector.showDialog oldColor, callingControl
    
    choosePDColor = dialog_ColorSelector.DialogResult
    If choosePDColor = vbOK Then newColor = dialog_ColorSelector.newColor
    
    Unload dialog_ColorSelector
    Set dialog_ColorSelector = Nothing

End Function

'Present a dialog box to confirm the closing of an unsaved image
Public Function confirmClose(ByVal formID As Long) As VbMsgBoxResult

    Load dialog_UnsavedChanges
    
    dialog_UnsavedChanges.formID = formID
    dialog_UnsavedChanges.showDialog FormMain
    
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
Public Function promptJP2Settings(ByRef srcImage As pdImage) As VbMsgBoxResult

    Load dialog_ExportJP2
    Set dialog_ExportJP2.imageBeingExported = srcImage
    dialog_ExportJP2.showDialog

    promptJP2Settings = dialog_ExportJP2.DialogResult
    
    Set dialog_ExportJP2.imageBeingExported = Nothing
    
    Unload dialog_ExportJP2
    Set dialog_ExportJP2 = Nothing

End Function

'Present a dialog box to ask the user for various WebP export settings
Public Function promptWebPSettings(ByRef srcImage As pdImage) As VbMsgBoxResult

    Load dialog_ExportWebP
    Set dialog_ExportWebP.imageBeingExported = srcImage
    dialog_ExportWebP.showDialog

    promptWebPSettings = dialog_ExportWebP.DialogResult
    
    Set dialog_ExportWebP.imageBeingExported = Nothing
    
    Unload dialog_ExportWebP
    Set dialog_ExportWebP = Nothing

End Function

'Present a dialog box to ask the user for various JPEG XR export settings
Public Function promptJXRSettings(ByRef srcImage As pdImage) As VbMsgBoxResult

    Load dialog_ExportJXR
    Set dialog_ExportJXR.imageBeingExported = srcImage
    dialog_ExportJXR.showDialog

    promptJXRSettings = dialog_ExportJXR.DialogResult
    
    Set dialog_ExportJXR.imageBeingExported = Nothing
    
    Unload dialog_ExportJXR
    Set dialog_ExportJXR = Nothing

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
Public Function promptAlphaCutoff(ByRef srcDIB As pdDIB) As VbMsgBoxResult

    Load dialog_AlphaCutoff
    dialog_AlphaCutoff.refDIB = srcDIB
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

'If an unclean shutdown + old Autosave data is found, offer to restore it for the user.
Public Function displayAutosaveWarning(ByRef dstArray() As autosaveXML) As VbMsgBoxResult

    Load dialog_AutosaveWarning
    dialog_AutosaveWarning.showDialog
    
    displayAutosaveWarning = dialog_AutosaveWarning.DialogResult
    
    'It's a bit unorthodox, but we must also populate dstArray() from this function, rather than relying on the
    ' dialog itself to do it (as VB makes it difficult to pass module-level array references).
    dialog_AutosaveWarning.fillArrayWithSaveResults dstArray
    
    Unload dialog_AutosaveWarning
    Set dialog_AutosaveWarning = Nothing

End Function

'A thin wrapper to showPDDialog, customized for resizing.
Public Sub showResizeDialog(ByVal ResizeTarget As PD_ACTION_TARGET)

    'Notify the resize dialog of the intended target
    FormResize.ResizeTarget = ResizeTarget

    'Display the resize dialog
    showPDDialog vbModal, FormResize

End Sub

'A thin wrapper to showPDDialog, customized for arbitrary rotation.
Public Sub showRotateDialog(ByVal RotateTarget As PD_ACTION_TARGET)

    'Notify the resize dialog of the intended target
    FormRotate.RotateTarget = RotateTarget

    'Display the resize dialog
    showPDDialog vbModal, FormRotate

End Sub

'A thin wrapper to showPDDialog, customized for arbitrary rotation.
Public Sub showStraightenDialog(ByVal StraightenTarget As PD_ACTION_TARGET)

    'Notify the resize dialog of the intended target
    FormStraighten.StraightenTarget = StraightenTarget

    'Display the resize dialog
    showPDDialog vbModal, FormStraighten

End Sub

