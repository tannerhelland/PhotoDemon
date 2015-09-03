Attribute VB_Name = "Dialog_Handler"
'***************************************************************************
'Custom Dialog Interface
'Copyright 2012-2015 by Tanner Helland
'Created: 30/November/12
'Last updated: 04/May/15
'Last update: start work on a generic "remember my choice" dialog, which will greatly simplify future tasks
'
'Module for handling all custom dialog forms used by PhotoDemon.  There are quite a few already, and I expect
' the number to grow as I phase out generic message boxes in favor of more descriptive (and usable) dialogs
' designed around a specific purpose.
'
'All dialogs are based off the same template, as you can see - they are just modal forms with a specially
' designed ".ShowDialog" sub or function that sets a ".DialogResult" property.  The wrapper function in this
' module simply checks that value, unloads the dialog form, then returns the value; this keeps all load/unload
' burdens here so that calling functions can simply use a MsgBox-style line to call custom dialogs and retrieve
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
Public Function displayAutosaveWarning(ByRef dstArray() As AutosaveXML) As VbMsgBoxResult

    Load dialog_AutosaveWarning
    dialog_AutosaveWarning.showDialog
    
    displayAutosaveWarning = dialog_AutosaveWarning.DialogResult
    
    'It's a bit unorthodox, but we must also populate dstArray() from this function, rather than relying on the
    ' dialog itself to do it (as VB makes it difficult to pass module-level array references).
    dialog_AutosaveWarning.fillArrayWithSaveResults dstArray
    
    Unload dialog_AutosaveWarning
    Set dialog_AutosaveWarning = Nothing

End Function

'A thin wrapper to showPDDialog, customized for generic resizing.
Public Sub showResizeDialog(ByVal ResizeTarget As PD_ACTION_TARGET)

    'Notify the resize dialog of the intended target
    FormResize.ResizeTarget = ResizeTarget

    'Display the resize dialog
    showPDDialog vbModal, FormResize

End Sub

'A thin wrapper to showPDDialog, customized for content-aware resizing.
Public Sub showContentAwareResizeDialog(ByVal ResizeTarget As PD_ACTION_TARGET)

    'Notify the resize dialog of the intended target
    FormResizeContentAware.ResizeTarget = ResizeTarget

    'Display the resize dialog
    showPDDialog vbModal, FormResizeContentAware

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

'Present a dialog box to ask the user how they want to tone map an incoming high bit-depth image.  Unlike other dialog
' requests, this one returns a pdParamString.  This is necessary because the return may have multiple parameters.
Public Function promptToneMapSettings(ByVal fi_Handle As Long, ByRef copyOfParamString As String) As VbMsgBoxResult
    
    'Before displaying the dialog, see if the user has requested that we automatically display previously specified settings
    If g_UserPreferences.GetPref_Boolean("Loading", "Tone Mapping Prompt", True) Then
    
        'Load the dialog, and supply it with any information it needs prior to display
        Load dialog_ToneMapping
        dialog_ToneMapping.fi_HandleCopy = fi_Handle
        
        'Display the (modal) dialog and wait for it to return
        dialog_ToneMapping.showDialog
        
        'This function will return the actual dialog result (OK vs Cancel)...
        promptToneMapSettings = dialog_ToneMapping.DialogResult
        
        If promptToneMapSettings = vbOK Then
        
            '...but we also need to return a copy of the parameter string, which FreeImage will use to actually render
            ' any requested tone-mapping operations.
            copyOfParamString = dialog_ToneMapping.toneMapSettings
            
            'If the user doesn't want us to raise this dialog in the future, store their preference now
            g_UserPreferences.SetPref_Boolean "Loading", "Tone Mapping Prompt", Not dialog_ToneMapping.RememberSettings
            
            'Write the param string out to the preferences file (in case the user decides to toggle this preference
            ' from the preferences dialog, or if they want settings automatically applied going forward).
            g_UserPreferences.SetPref_String "Loading", "Tone Mapping Settings", copyOfParamString
            
        End If
            
        'Release any other references, then exit
        Unload dialog_ToneMapping
        Set dialog_ToneMapping = Nothing
        
    'The user has requested that we do not prompt them for tone-map settings.  Use whatever settings they have
    ' previously specified.  If no settings were previously specified (meaning they disabled this preference prior
    ' to actually loading an HDR image, argh), generate a default set of "good enough" parameters.
    Else
    
        copyOfParamString = g_UserPreferences.GetPref_String("Loading", "Tone Mapping Settings", "")
        
        'Check for an empty string; if found, build a default param string
        If Len(copyOfParamString) = 0 Then
            copyOfParamString = buildParams(1, 0, 0)
        End If
        
        'Return "OK"
        promptToneMapSettings = vbOK
    
    End If

End Function

'Present an "add new preset" dialog box to the user.
Public Function promptNewPreset(ByRef srcPresetManager As pdToolPreset, ByRef parentForm As Form, ByRef dstPresetName As String) As VbMsgBoxResult

    Load dialog_AddPreset
    dialog_AddPreset.showDialog srcPresetManager, parentForm

    promptNewPreset = dialog_AddPreset.DialogResult
    
    dstPresetName = dialog_AddPreset.newPresetName
    
    Unload dialog_AddPreset
    Set dialog_AddPreset = Nothing

End Function

'Present a generic Yes/No dialog with an option to remember the current setting.  Once the option to remember has been set,
' it cannot be unset short of using the Reset button in the Tools > Options panel.
'
'The caller must supply a unique "questionID" string.  This is the string used to identify this dialog in the XML file,
' so it will be forced to an XML-safe equivalent.  As such, do not do something stupid like having two IDs that are so similar,
' their XML-safe variants become identical.
'
'Prompt text, "yes button" text, "no button" text, "cancel button" text, and icon (message box style) must be passed.
' The bottom "Remember my decision" text is universal and cannot be changed by the caller.
'
'If the user has previously ticked the "remember my decision" box, this function should still be called, but it will simply
' retrieve the previous choice and silently return it.
'
'Returns a VbMsgBoxResult constant, with YES, NO, or CANCEL specified.
Public Function promptGenericYesNoDialog(ByVal questionID As String, ByVal questionText As String, ByVal yesButtonText As String, ByVal noButtonText As String, ByVal cancelButtonText As String, ByVal rememberCheckBoxText As String, ByVal dialogTitleText As String, Optional ByVal icon As SystemIconConstants = IDI_QUESTION, Optional ByVal defaultAnswer As VbMsgBoxResult = vbCancel, Optional ByVal defaultRemember As Boolean = False) As VbMsgBoxResult

    'Convert the questionID to its XML-safe equivalent
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    questionID = xmlEngine.getXMLSafeTagName(questionID)
    
    'See if the user has already answered this question in the past.
    If g_UserPreferences.doesValueExist("Dialogs", questionID) Then
        
        'The user has already answered this question and saved their answer.  Retrieve the previous answer and exit.
        promptGenericYesNoDialog = g_UserPreferences.GetPref_Long("Dialogs", questionID, defaultAnswer)
        
    'The user has not saved a previous answer.  Display the full dialog.
    Else
    
        dialog_GenericMemory.showDialog questionText, yesButtonText, noButtonText, cancelButtonText, rememberCheckBoxText, dialogTitleText, icon, defaultAnswer, defaultRemember
        
        'Retrieve the user's answer
        promptGenericYesNoDialog = dialog_GenericMemory.DialogResult
        
        'If the user wants us to permanently remember this action, save their preference now.
        If dialog_GenericMemory.getRememberAnswerState Then
            g_UserPreferences.WritePreference "Dialogs", questionID, Trim$(Str(promptGenericYesNoDialog))
        End If
        
        'Release the dialog form
        Unload dialog_GenericMemory
        Set dialog_GenericMemory = Nothing
    
    End If

End Function

'Identical to promptGenericYesNoDialog(), above, with the caveat that only ONE possible outcome can be remembered.  This is relevant for
' Yes/No/Cancel situations where No and Cancel prevent a workflow from proceeding.  If we allowed those values to be stored, the user
' could never proceed with an operation in the future!
Public Function promptGenericYesNoDialog_SingleOutcome(ByVal questionID As String, ByVal questionText As String, ByVal yesButtonText As String, ByVal noButtonText As String, ByVal cancelButtonText As String, ByVal rememberCheckBoxText As String, ByVal dialogTitleText As String, Optional ByVal choiceAllowedToRemember As VbMsgBoxResult = vbYes, Optional ByVal icon As SystemIconConstants = IDI_QUESTION, Optional ByVal defaultAnswer As VbMsgBoxResult = vbCancel, Optional ByVal defaultRemember As Boolean = False) As VbMsgBoxResult

    'Convert the questionID to its XML-safe equivalent
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    questionID = xmlEngine.getXMLSafeTagName(questionID)
    
    'See if the user has already answered this question in the past.
    If g_UserPreferences.doesValueExist("Dialogs", questionID) Then
        
        'The user has already answered this question and saved their answer.  Retrieve the previous answer and exit.
        promptGenericYesNoDialog_SingleOutcome = g_UserPreferences.GetPref_Long("Dialogs", questionID, defaultAnswer)
        
    'The user has not saved a previous answer.  Display the full dialog.
    Else
    
        dialog_GenericMemory.showDialog questionText, yesButtonText, noButtonText, cancelButtonText, rememberCheckBoxText, dialogTitleText, icon, defaultAnswer, defaultRemember
        
        'Retrieve the user's answer
        promptGenericYesNoDialog_SingleOutcome = dialog_GenericMemory.DialogResult
        
        'If the user wants us to permanently remember this action, save their preference now.
        If dialog_GenericMemory.getRememberAnswerState Then
            g_UserPreferences.WritePreference "Dialogs", questionID, Trim$(Str(choiceAllowedToRemember))
        End If
        
        'Release the dialog form
        Unload dialog_GenericMemory
        Set dialog_GenericMemory = Nothing
    
    End If

End Function

'Present the user with PD's custom brush selection dialog.
' INPUTS:  1) a String-type variable (ByRef, of course) which will receive the new brush parameters
'          2) an optional initial brush parameter string
'          3) an optional brushSelector control reference, if this dialog is being raised by a brushSelector control.
'             (This reference will be used to provide live updates as the user plays with the brush dialog.)
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function showBrushDialog(ByRef newBrush As String, Optional ByVal initialBrush As String = "", Optional ByRef callingControl As brushSelector) As Boolean
    
    If choosePDBrush(initialBrush, newBrush, callingControl) = vbOK Then
        showBrushDialog = True
    Else
        showBrushDialog = False
    End If
    
End Function

'Display a custom brush selection dialog
Public Function choosePDBrush(ByRef oldBrush As String, ByRef newBrush As String, Optional ByRef callingControl As brushSelector) As VbMsgBoxResult

    Load dialog_FillSettings
    dialog_FillSettings.showDialog oldBrush, callingControl
    
    choosePDBrush = dialog_FillSettings.DialogResult
    If choosePDBrush = vbOK Then newBrush = dialog_FillSettings.newBrush
    
    Unload dialog_FillSettings
    Set dialog_FillSettings = Nothing

End Function

'Present the user with PD's custom pen selection dialog.
' INPUTS:  1) a String-type variable (ByRef, of course) which will receive the new pen parameters
'          2) an optional initial pen parameter string
'          3) an optional penSelector control reference, if this dialog is being raised by a penSelector control.
'             (This reference will be used to provide live updates as the user plays with the pen dialog.)
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function showPenDialog(ByRef newPen As String, Optional ByVal initialPen As String = "", Optional ByRef callingControl As penSelector) As Boolean
    
    If choosePDPen(initialPen, newPen, callingControl) = vbOK Then
        showPenDialog = True
    Else
        showPenDialog = False
    End If
    
End Function

'Display a custom pen selection dialog
Public Function choosePDPen(ByRef oldPen As String, ByRef newPen As String, Optional ByRef callingControl As penSelector) As VbMsgBoxResult

    Load dialog_OutlineSettings
    dialog_OutlineSettings.showDialog oldPen, callingControl
    
    choosePDPen = dialog_OutlineSettings.DialogResult
    If choosePDPen = vbOK Then newPen = dialog_OutlineSettings.newPen
    
    Unload dialog_OutlineSettings
    Set dialog_OutlineSettings = Nothing

End Function

'Present the user with PD's custom gradient selection dialog.
' INPUTS:  1) a String-type variable (ByRef, of course) which will receive the new gradient parameters
'          2) an optional initial gradient parameter string
'          3) an optional gradientSelector control reference, if this dialog is being raised by a gradientSelector control.
'             (This reference will be used to provide live updates as the user plays with the dialog.)
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function showGradientDialog(ByRef newGradient As String, Optional ByVal initialGradient As String = "", Optional ByRef callingControl As gradientSelector) As Boolean
    
    If choosePDGradient(initialGradient, newGradient, callingControl) = vbOK Then
        showGradientDialog = True
    Else
        showGradientDialog = False
    End If
    
End Function

'Display a custom gradient selection dialog
' RETURNS: MsgBoxResult from the dialog itself.  For easier interactions, I recommend using the showGradientDialog function, above.
Public Function choosePDGradient(ByRef oldGradient As String, ByRef newGradient As String, Optional ByRef callingControl As gradientSelector) As VbMsgBoxResult

    Load dialog_GradientEditor
    dialog_GradientEditor.showDialog oldGradient, callingControl
    
    choosePDGradient = dialog_GradientEditor.DialogResult
    If choosePDGradient = vbOK Then newGradient = dialog_GradientEditor.newGradient
    
    Unload dialog_GradientEditor
    Set dialog_GradientEditor = Nothing

End Function

