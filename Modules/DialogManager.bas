Attribute VB_Name = "Dialogs"
'***************************************************************************
'Custom Dialog Interface
'Copyright 2012-2026 by Tanner Helland
'Created: 30/November/12
'Last updated: 01/March/24
'Last update: small wrapper for PDF import settings
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Display a custom color selection dialog
Public Function ChoosePDColor(ByVal oldColor As Long, ByRef newColor As Long, Optional ByRef callingControl As pdColorSelector, Optional ByRef callerParent As Form = Nothing) As VbMsgBoxResult
    Load dialog_ColorSelector
    dialog_ColorSelector.ShowDialog oldColor, callingControl, callerParent
    ChoosePDColor = dialog_ColorSelector.DialogResult
    If (ChoosePDColor = vbOK) Then newColor = dialog_ColorSelector.NewlySelectedColor
    Unload dialog_ColorSelector
    Set dialog_ColorSelector = Nothing
End Function

'Present a dialog box to confirm the closing of an unsaved image
Public Function ConfirmClose(ByVal srcImageID As Long, ByVal numOfUnsavedImages As Long, ByRef unsavedImageIDs As pdStack) As VbMsgBoxResult
    Load dialog_UnsavedChanges
    dialog_UnsavedChanges.ShowDialog srcImageID, numOfUnsavedImages, unsavedImageIDs
    ConfirmClose = dialog_UnsavedChanges.DialogResult
    Unload dialog_UnsavedChanges
    Set dialog_UnsavedChanges = Nothing
End Function

Public Function PromptAVIFSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportAVIF
    dialog_ExportAVIF.ShowDialog srcImage
    
    PromptAVIFSettings = dialog_ExportAVIF.GetDialogResult
    dstFormatParams = dialog_ExportAVIF.GetFormatParams
    
    'AVIF metadata isn't currently supported (via ExifTool; we could probably handle it manually by
    ' writing a separate XMP file and embedding it via avifenc, but I haven't explored this thoroughly).
    ' If this changes, the following line can be changed to match any new metadata features
    'dstMetadataParams = dialog_ExportAVIF.GetMetadataParams
    dstMetadataParams = vbNullString
    
    Unload dialog_ExportAVIF
    Set dialog_ExportAVIF = Nothing
    
End Function

Public Function PromptBMPSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportBMP
    dialog_ExportBMP.ShowDialog srcImage

    PromptBMPSettings = dialog_ExportBMP.GetDialogResult
    dstFormatParams = dialog_ExportBMP.GetFormatParams
    
    'The BMP format does not currently support metadata, but if it ever does, this line can be changed to match
    dstMetadataParams = vbNullString        'dialog_ExportBMP.GetMetadataParams
    
    Unload dialog_ExportBMP
    Set dialog_ExportBMP = Nothing
    
End Function

Public Function PromptDDSSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportDDS
    dialog_ExportDDS.ShowDialog srcImage
    
    PromptDDSSettings = dialog_ExportDDS.GetDialogResult
    dstFormatParams = dialog_ExportDDS.GetFormatParams
    
    'DDS metadata isn't currently supported (via ExifTool).
    ' If this changes, the following line can be changed to match any new metadata features
    'dstMetadataParams = dialog_ExportDDS.GetMetadataParams
    dstMetadataParams = vbNullString
    
    Unload dialog_ExportDDS
    Set dialog_ExportDDS = Nothing
    
End Function

Public Function PromptGIFSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportGIF
    dialog_ExportGIF.ShowDialog srcImage
    
    PromptGIFSettings = dialog_ExportGIF.GetDialogResult
    dstFormatParams = dialog_ExportGIF.GetFormatParams
    dstMetadataParams = dialog_ExportGIF.GetMetadataParams
    
    Unload dialog_ExportGIF
    Set dialog_ExportGIF = Nothing
    
End Function

Public Function PromptHEIFSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportHEIF
    dialog_ExportHEIF.ShowDialog srcImage
    
    PromptHEIFSettings = dialog_ExportHEIF.GetDialogResult
    dstFormatParams = dialog_ExportHEIF.GetFormatParams
    dstMetadataParams = dialog_ExportHEIF.GetMetadataParams
    
    Unload dialog_ExportHEIF
    Set dialog_ExportHEIF = Nothing
    
End Function

Public Function PromptICOSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportICO
    dialog_ExportICO.ShowDialog srcImage
    
    PromptICOSettings = dialog_ExportICO.GetDialogResult
    dstFormatParams = dialog_ExportICO.GetFormatParams
    dstMetadataParams = dialog_ExportICO.GetMetadataParams
    
    Unload dialog_ExportICO
    Set dialog_ExportICO = Nothing
    
End Function

Public Function PromptJP2Settings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult

    Load dialog_ExportJP2
    dialog_ExportJP2.ShowDialog srcImage
    
    PromptJP2Settings = dialog_ExportJP2.GetDialogResult
    dstFormatParams = dialog_ExportJP2.GetFormatParams
    dstMetadataParams = dialog_ExportJP2.GetMetadataParams
    
    Unload dialog_ExportJP2
    Set dialog_ExportJP2 = Nothing
    
End Function

Public Function PromptJPEGSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportJPEG
    dialog_ExportJPEG.ShowDialog srcImage
    
    PromptJPEGSettings = dialog_ExportJPEG.GetDialogResult
    dstFormatParams = dialog_ExportJPEG.GetFormatParams
    dstMetadataParams = dialog_ExportJPEG.GetMetadataParams
    
    Unload dialog_ExportJPEG
    Set dialog_ExportJPEG = Nothing
    
End Function

Public Function PromptJXLSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportJXL
    dialog_ExportJXL.ShowDialog srcImage
    
    PromptJXLSettings = dialog_ExportJXL.GetDialogResult
    dstFormatParams = dialog_ExportJXL.GetFormatParams
    dstMetadataParams = dialog_ExportJXL.GetMetadataParams
    
    Unload dialog_ExportJXL
    Set dialog_ExportJXL = Nothing
    
End Function

Public Function PromptJXRSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult

    Load dialog_ExportJXR
    dialog_ExportJXR.ShowDialog srcImage
    
    PromptJXRSettings = dialog_ExportJXR.GetDialogResult
    dstFormatParams = dialog_ExportJXR.GetFormatParams
    dstMetadataParams = dialog_ExportJXR.GetMetadataParams
    
    Unload dialog_ExportJXR
    Set dialog_ExportJXR = Nothing
    
End Function

Public Function PromptPaletteSettings(ByRef srcImage As pdImage, ByVal palFormat As PD_PaletteFormat, ByRef dstFilename As String, ByRef dstFormatParams As String) As VbMsgBoxResult
    
    Load dialog_ExportPalette
    dialog_ExportPalette.ShowDialog srcImage, palFormat, dstFilename
    
    PromptPaletteSettings = dialog_ExportPalette.GetDialogResult
    dstFormatParams = dialog_ExportPalette.GetFormatParams
    
    Unload dialog_ExportPalette
    Set dialog_ExportPalette = Nothing
    
End Function

Public Function PromptPNGSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportPNG
    dialog_ExportPNG.ShowDialog srcImage
    
    PromptPNGSettings = dialog_ExportPNG.GetDialogResult
    dstFormatParams = dialog_ExportPNG.GetFormatParams
    dstMetadataParams = dialog_ExportPNG.GetMetadataParams
    
    Unload dialog_ExportPNG
    Set dialog_ExportPNG = Nothing
    
End Function

Public Function PromptPNMSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportPixmap
    dialog_ExportPixmap.ShowDialog srcImage
    
    PromptPNMSettings = dialog_ExportPixmap.GetDialogResult
    dstFormatParams = dialog_ExportPixmap.GetFormatParams
    dstMetadataParams = dialog_ExportPixmap.GetMetadataParams
    
    Unload dialog_ExportPixmap
    Set dialog_ExportPixmap = Nothing
    
End Function

Public Function PromptPSDSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportPSD
    dialog_ExportPSD.ShowDialog srcImage
    
    PromptPSDSettings = dialog_ExportPSD.GetDialogResult
    dstFormatParams = dialog_ExportPSD.GetFormatParams
    dstMetadataParams = dialog_ExportPSD.GetMetadataParams
    
    Unload dialog_ExportPSD
    Set dialog_ExportPSD = Nothing
    
End Function

Public Function PromptPSPSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportPSP
    dialog_ExportPSP.ShowDialog srcImage
    
    PromptPSPSettings = dialog_ExportPSP.GetDialogResult
    dstFormatParams = dialog_ExportPSP.GetFormatParams
    dstMetadataParams = dialog_ExportPSP.GetMetadataParams
    
    Unload dialog_ExportPSP
    Set dialog_ExportPSP = Nothing
    
End Function

Public Function PromptTIFFSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportTIFF
    dialog_ExportTIFF.ShowDialog srcImage
    
    PromptTIFFSettings = dialog_ExportTIFF.GetDialogResult
    dstFormatParams = dialog_ExportTIFF.GetFormatParams
    dstMetadataParams = dialog_ExportTIFF.GetMetadataParams
    
    Unload dialog_ExportTIFF
    Set dialog_ExportTIFF = Nothing
    
End Function

Public Function PromptWebPSettings(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult

    Load dialog_ExportWebP
    dialog_ExportWebP.ShowDialog srcImage
    
    PromptWebPSettings = dialog_ExportWebP.GetDialogResult
    dstFormatParams = dialog_ExportWebP.GetFormatParams
    dstMetadataParams = dialog_ExportWebP.GetMetadataParams
    
    Unload dialog_ExportWebP
    Set dialog_ExportWebP = Nothing
    
End Function

'If the user is running in the IDE, warn them of the consequences of doing so
Public Function DisplayIDEWarning() As VbMsgBoxResult

    Load dialog_IDEWarning
    dialog_IDEWarning.ShowDialog

    DisplayIDEWarning = dialog_IDEWarning.DialogResult
    
    Unload dialog_IDEWarning
    Set dialog_IDEWarning = Nothing

End Function

'If an unclean shutdown + old Autosave data is found, offer to restore it for the user.
Public Function DisplayAutosaveWarning(ByRef dstArray() As AutosaveXML) As VbMsgBoxResult

    Load dialog_AutosaveWarning
    dialog_AutosaveWarning.ShowDialog
    
    DisplayAutosaveWarning = dialog_AutosaveWarning.DialogResult
    
    'It's a bit unorthodox, but we must also populate dstArray() from this function, rather than relying on the
    ' dialog itself to do it (as VB makes it difficult to pass module-level array references).
    dialog_AutosaveWarning.FillArrayWithSaveResults dstArray
    
    Unload dialog_AutosaveWarning
    Set dialog_AutosaveWarning = Nothing

End Function

'A thin wrapper to showPDDialog, customized for generic resizing.
Public Sub ShowResizeDialog(ByVal ResizeTarget As PD_ActionTarget)
    FormResize.ResizeTarget = ResizeTarget
    ShowPDDialog vbModal, FormResize
End Sub

'A thin wrapper to showPDDialog, customized for content-aware resizing.
Public Sub ShowContentAwareResizeDialog(ByVal ResizeTarget As PD_ActionTarget)
    FormResizeContentAware.ResizeTarget = ResizeTarget
    ShowPDDialog vbModal, FormResizeContentAware
End Sub

'A thin wrapper to showPDDialog, customized for arbitrary rotation.
Public Sub ShowRotateDialog(ByVal RotateTarget As PD_ActionTarget)
    FormRotate.RotateTarget = RotateTarget
    ShowPDDialog vbModal, FormRotate
End Sub

'A thin wrapper to showPDDialog, customized for arbitrary rotation.
Public Sub ShowStraightenDialog(ByVal StraightenTarget As PD_ActionTarget)
    FormStraighten.StraightenTarget = StraightenTarget
    ShowPDDialog vbModal, FormStraighten
End Sub

'Present a dialog box to ask the user how they want to tone map an incoming high bit-depth image.  Unlike other dialog
' requests, this one returns an XML string.  This is necessary because the return may have multiple parameters.
Public Function PromptToneMapSettings(ByVal fi_Handle As Long, ByRef copyOfParamString As String) As VbMsgBoxResult
    
    'Before displaying the dialog, see if the user has requested that we automatically display previously specified settings
    If UserPrefs.GetPref_Boolean("Loading", "Tone Mapping Prompt", True) Then
    
        'Load the dialog, and supply it with any information it needs prior to display
        Load dialog_ToneMapping
        dialog_ToneMapping.fi_HandleCopy = fi_Handle
        
        'Display the (modal) dialog and wait for it to return
        dialog_ToneMapping.ShowDialog
        
        'This function will return the actual dialog result (OK vs Cancel)...
        PromptToneMapSettings = dialog_ToneMapping.DialogResult
        
        If (PromptToneMapSettings = vbOK) Then
        
            '...but we also need to return a copy of the parameter string, which FreeImage will use to actually render
            ' any requested tone-mapping operations.
            copyOfParamString = dialog_ToneMapping.ToneMapSettings
            
            'If the user doesn't want us to raise this dialog in the future, store their preference now
            UserPrefs.SetPref_Boolean "Loading", "Tone Mapping Prompt", Not dialog_ToneMapping.RememberSettings
            
            'Write the param string out to the preferences file (in case the user decides to toggle this preference
            ' from the preferences dialog, or if they want settings automatically applied going forward).
            UserPrefs.SetPref_String "Loading", "Tone Mapping Settings", copyOfParamString
            
        End If
            
        'Release any other references, then exit
        Unload dialog_ToneMapping
        Set dialog_ToneMapping = Nothing
        
    'The user has requested that we do not prompt them for tone-map settings.  Use whatever settings they have
    ' previously specified.  If no settings were previously specified (meaning they disabled this preference prior
    ' to actually loading an HDR image, argh), generate a default set of "good enough" parameters.
    Else
    
        copyOfParamString = UserPrefs.GetPref_String("Loading", "Tone Mapping Settings", vbNullString)
        
        'Check for an empty string; if found, build a default param string
        If (LenB(copyOfParamString) = 0) Then copyOfParamString = BuildParamList("method", PDTM_DRAGO)
        
        'Return "OK"
        PromptToneMapSettings = vbOK
    
    End If

End Function

'Present a generic input box (single edit box for user entry).
' Text passed to this function *should already be localized*.
Public Function PromptInputBox(ByRef dlgTitle As String, ByRef dlgPrompt As String, ByRef inOutInputText As String) As VbMsgBoxResult

    Load dialog_InputBox
    Interface.FixPopupWindow dialog_InputBox.hWnd, True
    dialog_InputBox.ShowDialog dlgTitle, dlgPrompt, inOutInputText
    
    PromptInputBox = dialog_InputBox.DialogResult
    inOutInputText = dialog_InputBox.UserEnteredText
    Interface.FixPopupWindow dialog_InputBox.hWnd, False
    
    Unload dialog_InputBox
    Set dialog_InputBox = Nothing

End Function

'Present an "add new preset" dialog box to the user.
Public Function PromptNewPreset(ByRef srcPresetManager As pdToolPreset, ByRef srcCommandBar As pdCommandBar, ByRef parentForm As Form) As VbMsgBoxResult

    Load dialog_AddPreset
    Interface.FixPopupWindow dialog_AddPreset.hWnd, True
    dialog_AddPreset.ShowDialog srcPresetManager, srcCommandBar, parentForm
    
    PromptNewPreset = dialog_AddPreset.DialogResult
    Interface.FixPopupWindow dialog_AddPreset.hWnd, False
    
    Unload dialog_AddPreset
    Set dialog_AddPreset = Nothing

End Function

'Present a generic binary choice dialog (typically Yes/No/Cancel) with an option to remember the current setting.
' Once the option to remember has been set, it can only be unset by:
' 1) Tools > Options > Interface > Reset all "remember my choice" decisions
' 2) Tools > Options > Advanced > Reset all preferences
'
'The caller must supply a unique "questionID" string.  This is the string used to identify this dialog in the XML file.
' This value will be forced to an XML-safe equivalent.  As such, do not do something stupid like having two IDs that vary
' only by whitespace (so their XML-safe variants would be identical).
'
'Prompt text, "yes button" text, "no button" text, "cancel button" text, "remember this decision" text, and icon
' (message box style) must be passed.
'
'If the user has previously ticked the "remember my decision" box, this function should still be called,
' but it will simply query the previously remembered choice and silently return value.
'
'Returns a VbMsgBoxResult constant, with YES, NO, or CANCEL specified.
Public Function PromptGenericYesNoDialog(ByVal questionID As String, ByRef questionText As String, ByRef yesButtonText As String, ByRef noButtonText As String, ByRef cancelButtonText As String, ByRef rememberCheckBoxText As String, ByRef dialogTitleText As String, Optional ByVal sysIcon As SystemIconConstants = 0, Optional ByVal defaultAnswer As VbMsgBoxResult = vbCancel, Optional ByVal defaultRemember As Boolean = False, Optional ByVal resNameYesImg As String = "generic_ok", Optional ByVal resNameNoImg As String = "generic_cancel", Optional ByVal resNameCancelImg As String = vbNullString) As VbMsgBoxResult

    'Convert the questionID to its XML-safe equivalent
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    questionID = xmlEngine.GetXMLSafeTagName(questionID)
    
    'See if the user has already answered this question in the past.
    If UserPrefs.DoesValueExist("Dialogs", questionID) Then
        
        'The user has already answered this question and saved their answer.  Retrieve the previous answer and exit.
        PromptGenericYesNoDialog = UserPrefs.GetPref_Long("Dialogs", questionID, defaultAnswer)
        
    'The user has not saved a previous answer.  Display the full dialog.
    Else
        
        dialog_GenericMemory.ShowDialog questionText, yesButtonText, noButtonText, cancelButtonText, rememberCheckBoxText, dialogTitleText, sysIcon, defaultAnswer, defaultRemember, resNameYesImg, resNameNoImg, resNameCancelImg
        
        'Retrieve the user's answer
        PromptGenericYesNoDialog = dialog_GenericMemory.DialogResult
        
        'If the user wants us to permanently remember this action, save their preference now.
        If dialog_GenericMemory.GetRememberAnswerState Then UserPrefs.WritePreference "Dialogs", questionID, Trim$(Str$(PromptGenericYesNoDialog))
        
        Unload dialog_GenericMemory
        Set dialog_GenericMemory = Nothing
    
    End If

End Function

'Identical to PromptGenericYesNoDialog(), above, with the caveat that only ONE possible outcome (typically YES)
' is allowed to be remembered.
'
'This is relevant for Yes/No/Cancel situations where No and Cancel prevent a workflow from proceeding.
' If we allowed No/Cancel values to be stored, the user could never proceed with an operation in the future!
Public Function PromptGenericYesNoDialog_SingleOutcome(ByVal questionID As String, ByRef questionText As String, ByRef yesButtonText As String, ByRef noButtonText As String, ByRef cancelButtonText As String, ByRef rememberCheckBoxText As String, ByRef dialogTitleText As String, Optional ByVal choiceAllowedToRemember As VbMsgBoxResult = vbYes, Optional ByVal useIcon As SystemIconConstants = IDI_QUESTION, Optional ByVal defaultAnswer As VbMsgBoxResult = vbCancel, Optional ByVal defaultRemember As Boolean = False) As VbMsgBoxResult

    'Convert the questionID to its XML-safe equivalent
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    questionID = xmlEngine.GetXMLSafeTagName(questionID)
    
    'See if the user has already answered this question in the past.
    If UserPrefs.DoesValueExist("Dialogs", questionID) Then
        
        'The user has already answered this question and saved their answer.  Retrieve the previous answer and exit.
        PromptGenericYesNoDialog_SingleOutcome = UserPrefs.GetPref_Long("Dialogs", questionID, defaultAnswer)
        
    'The user has not saved a previous answer.  Display the full dialog.
    Else
    
        dialog_GenericMemory.ShowDialog questionText, yesButtonText, noButtonText, cancelButtonText, rememberCheckBoxText, dialogTitleText, useIcon, defaultAnswer, defaultRemember
        
        'Retrieve the user's answer
        PromptGenericYesNoDialog_SingleOutcome = dialog_GenericMemory.DialogResult
        
        'If the user wants us to permanently remember this action, save their preference now.
        If dialog_GenericMemory.GetRememberAnswerState Then UserPrefs.WritePreference "Dialogs", questionID, Trim$(Str$(choiceAllowedToRemember))
        
        'Release the dialog form
        Unload dialog_GenericMemory
        Set dialog_GenericMemory = Nothing
    
    End If

End Function

'Present the user with PD's custom brush selection dialog.
' INPUTS:  1) a String-type variable (ByRef, of course) which will receive the new brush parameters
'          2) an optional initial brush parameter string
'          3) an optional pdBrushSelector control reference, if this dialog is being raised by a pdBrushSelector control.
'             (This reference will be used to provide live updates as the user plays with the brush dialog.)
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function ShowBrushDialog(ByRef newBrush As String, Optional ByVal initialBrush As String = vbNullString, Optional ByRef callingControl As pdBrushSelector) As Boolean
    ShowBrushDialog = (ChoosePDBrush(initialBrush, newBrush, callingControl) = vbOK)
End Function

'Display a custom brush selection dialog
Public Function ChoosePDBrush(ByRef oldBrush As String, ByRef newBrush As String, Optional ByRef callingControl As pdBrushSelector) As VbMsgBoxResult

    Load dialog_FillSettings
    dialog_FillSettings.ShowDialog oldBrush, callingControl
    
    ChoosePDBrush = dialog_FillSettings.DialogResult
    If ChoosePDBrush = vbOK Then newBrush = dialog_FillSettings.newBrush
    
    Unload dialog_FillSettings
    Set dialog_FillSettings = Nothing

End Function

'Present the user with PD's custom pen selection dialog.
' INPUTS:  1) a String-type variable (ByRef, of course) which will receive the new pen parameters
'          2) an optional initial pen parameter string
'          3) an optional pdPenSelector control reference, if this dialog is being raised by a pdPenSelector control.
'             (This reference will be used to provide live updates as the user plays with the pen dialog.)
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function ShowPenDialog(ByRef newPen As String, Optional ByVal initialPen As String = vbNullString, Optional ByRef callingControl As pdPenSelector) As Boolean
    ShowPenDialog = (ChoosePDPen(initialPen, newPen, callingControl) = vbOK)
End Function

'Display a custom pen selection dialog
Public Function ChoosePDPen(ByRef oldPen As String, ByRef newPen As String, Optional ByRef callingControl As pdPenSelector) As VbMsgBoxResult

    Load dialog_OutlineSettings
    dialog_OutlineSettings.ShowDialog oldPen, callingControl
    
    ChoosePDPen = dialog_OutlineSettings.DialogResult
    If ChoosePDPen = vbOK Then newPen = dialog_OutlineSettings.GetNewPen
    
    Unload dialog_OutlineSettings
    Set dialog_OutlineSettings = Nothing

End Function

'Present the user with PD's custom gradient selection dialog.
' INPUTS:  1) a String-type variable (ByRef, of course) which will receive the new gradient parameters
'          2) an optional initial gradient parameter string
'          3) an optional pdGradientSelector control reference, if this dialog is being raised by a pdGradientSelector control.
'             (This reference will be used to provide live updates as the user plays with the dialog.)
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function ShowGradientDialog(ByRef NewGradient As String, Optional ByVal initialGradient As String = vbNullString, Optional ByRef callingControl As pdGradientSelector) As Boolean
    ShowGradientDialog = (ChoosePDGradient(initialGradient, NewGradient, callingControl) = vbOK)
End Function

'Display a custom gradient selection dialog
' RETURNS: MsgBoxResult from the dialog itself.  For easier interactions, I recommend using the showGradientDialog function, above.
Public Function ChoosePDGradient(ByRef oldGradient As String, ByRef NewGradient As String, Optional ByRef callingControl As pdGradientSelector) As VbMsgBoxResult

    Load dialog_GradientEditor
    dialog_GradientEditor.ShowDialog oldGradient, callingControl
    
    ChoosePDGradient = dialog_GradientEditor.DialogResult
    If ChoosePDGradient = vbOK Then NewGradient = dialog_GradientEditor.NewGradient
    
    Unload dialog_GradientEditor
    Set dialog_GradientEditor = Nothing

End Function

'Present a first-run dialog box that gives the user a choice of language and UI theme
Public Function PromptUITheme() As VbMsgBoxResult
    
    'Falsafe check
    If (g_Themer Is Nothing) Then Exit Function
    
    'Before displaying the dialog, cache the current language and theme settings.  If the user changes
    ' one or more of these via the dialog, we need to repaint the main form after the dialog closes.
    Dim backupLangIndex As Long
    backupLangIndex = g_Language.GetCurrentLanguageIndex
    
    Dim backupThemeClass As PD_THEME_CLASS, backupThemeAccent As PD_THEME_ACCENT, backupIconsMono As Boolean
    backupThemeClass = g_Themer.GetCurrentThemeClass
    backupThemeAccent = g_Themer.GetCurrentThemeAccent
    backupIconsMono = g_Themer.GetMonochromeIconSetting
    
    Dim newLangIndex As Long
    Dim newThemeClass As PD_THEME_CLASS, newThemeAccent As PD_THEME_ACCENT, newIconsMono As Boolean
    
    Load dialog_UITheme
    dialog_UITheme.ShowDialog
    PromptUITheme = dialog_UITheme.DialogResult
    
    'Retrieve any/all new settings from the dialog, then release it
    dialog_UITheme.GetNewSettings newLangIndex, newThemeClass, newThemeAccent, newIconsMono
    Unload dialog_UITheme
    Set dialog_UITheme = Nothing
    
    'Regardless of the return value, note that the user has seen this dialog
    UserPrefs.SetPref_Boolean "Themes", "HasSeenThemeDialog", True
    
    'If the dialog was canceled, reset the original language and theme.
    If (PromptUITheme <> vbOK) Then
        
        g_Language.ActivateNewLanguage backupLangIndex
        g_Language.ApplyLanguage False
        g_Themer.SetNewTheme backupThemeClass, backupThemeAccent
        g_Themer.SetMonochromeIconSetting backupIconsMono
    
    'If the dialog was *not* canceled, make the new settings persistent.  (Note that we only apply theme settings here,
    ' not language settings, as they are handled separately.)
    Else
        
        g_Themer.SetNewTheme newThemeClass, newThemeAccent
        g_Themer.SetMonochromeIconSetting newIconsMono
        
    End If
    
    'Four steps are required to activate a theme change:
    ' 1) Load the new theme (or accent) data from file
    ' 2) Notify the resource manager of the change (because things like UI icons may need to be redrawn)
    ' 3) Regenerate any/all internal rendering caches (some rely on theme colors)
    ' 4) Redraw the main window, including all child panels and controls
    g_Themer.LoadDefaultPDTheme
    g_Resources.NotifyThemeChange
    
    'If a new language is in use, apply it now
    If (PromptUITheme = vbOK) Then
        If (newLangIndex <> backupLangIndex) Then
            
            'Load the old language file and undo any existing translations
            g_Language.ActivateNewLanguage backupLangIndex
            g_Language.ApplyLanguage False
            
            g_Language.UndoTranslations FormMain
            g_Language.UndoTranslations toolbar_Toolbox
            g_Language.UndoTranslations toolbar_Options
            g_Language.UndoTranslations toolbar_Layers
            
            'Now, load the *new* language and apply it
            g_Language.ActivateNewLanguage newLangIndex
            g_Language.ApplyLanguage True
            
        End If
    End If
    
    'If the theme has actually changed, apply the changes now.  (We can skip this step if the dialog was canceled,
    ' or if the user confirmed their original settings.)
    If (PromptUITheme = vbOK) Then
        If (newThemeClass <> backupThemeClass) Or (newThemeAccent <> backupThemeAccent) Or (newIconsMono <> backupIconsMono) Then
            Drawing.CacheUIPensAndBrushes
            UserControls.NotifyTooltipThemeChange
            IconsAndCursors.LoadMenuIcons False
        End If
    End If
    
    'Finally, redraw the entire UI.  (Note that this happens *regardless* of dialog cancelation, as the dialog has
    ' to play some games with the theme engine to properly live-update itself - and we don't want any of those
    ' intermediate changes "filtering through" to the main UI.)
    Interface.RedrawEntireUI True
    
End Function

'Display the right-side color panel toolbox settings dialog
Public Function ChooseColorPanelSettings() As VbMsgBoxResult

    Load dialog_ColorPanel
    dialog_ColorPanel.ShowDialog
    
    ChooseColorPanelSettings = dialog_ColorPanel.DialogResult
    Unload dialog_ColorPanel
    Set dialog_ColorPanel = Nothing

End Function

'Ask the user if they want to drag/drop the target image as a new standalone image, or a new layer in the current image.
' (If no image is loaded, the drag/drop result will be forced to "drop as standalone image", e.g. this function will return NO.
'  Note that CANCEL is also a viable return, one which the user needs to respect.)
Public Function PromptDropAsNewLayer() As VbMsgBoxResult

    If (PDImages.GetNumOpenImages() = 0) Then
        PromptDropAsNewLayer = vbYes
    Else
        
        Dim questionText As String
        questionText = g_Language.TranslateMessage("How would you like to load this image?")
        
        Dim yesText As String, noText As String, cancelText As String
        yesText = g_Language.TranslateMessage("Open it as a standalone image.")
        noText = g_Language.TranslateMessage("Add it to the current image as a new layer.")
        cancelText = g_Language.TranslateMessage("I can't decide.  Cancel this action.")
        
        Dim rememberText As String, dialogTitle As String
        rememberText = g_Language.TranslateMessage("in the future, do this without asking me")
        dialogTitle = "Drag and drop options"
        
        'Display the dialog and return the result
        Dim questionID As String
        questionID = g_Language.TranslateMessage("Drop as image or layer")
        PromptDropAsNewLayer = Dialogs.PromptGenericYesNoDialog(questionID, questionText, yesText, noText, cancelText, rememberText, dialogTitle, IDI_QUESTION, vbCancel, False, "generic_image", "generic_add", "generic_cancel")
            
    End If

End Function

Public Sub PromptEffect_Animation(ByVal useBackgroundMode As Boolean)
    Load FormAnimBackground
    FormAnimBackground.SetBackgroundMode useBackgroundMode
    Interface.ShowPDDialog vbModal, FormAnimBackground
End Sub

Public Sub PromptEffect_Median(ByVal medianCutoff As Single)
    Load FormMedian
    FormMedian.SetMedianCutoff medianCutoff
    Interface.ShowPDDialog vbModal, FormMedian
End Sub

Public Function PromptExportAnimatedGIF(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportAnimatedGIF
    dialog_ExportAnimatedGIF.ShowDialog srcImage
    
    PromptExportAnimatedGIF = dialog_ExportAnimatedGIF.GetDialogResult
    dstFormatParams = dialog_ExportAnimatedGIF.GetFormatParams
    dstMetadataParams = dialog_ExportAnimatedGIF.GetMetadataParams
    
    Unload dialog_ExportAnimatedGIF
    Set dialog_ExportAnimatedGIF = Nothing
    
End Function

Public Function PromptExportAnimatedJXL(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportAnimatedJXL
    dialog_ExportAnimatedJXL.ShowDialog srcImage
    
    PromptExportAnimatedJXL = dialog_ExportAnimatedJXL.GetDialogResult
    dstFormatParams = dialog_ExportAnimatedJXL.GetFormatParams
    dstMetadataParams = dialog_ExportAnimatedJXL.GetMetadataParams
    
    Unload dialog_ExportAnimatedJXL
    Set dialog_ExportAnimatedJXL = Nothing
    
End Function

Public Function PromptExportAnimatedPNG(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportAnimatedPNG
    dialog_ExportAnimatedPNG.ShowDialog srcImage
    
    PromptExportAnimatedPNG = dialog_ExportAnimatedPNG.GetDialogResult
    dstFormatParams = dialog_ExportAnimatedPNG.GetFormatParams
    dstMetadataParams = dialog_ExportAnimatedPNG.GetMetadataParams
    
    Unload dialog_ExportAnimatedPNG
    Set dialog_ExportAnimatedPNG = Nothing
    
End Function

Public Function PromptExportAnimatedWebP(ByRef srcImage As pdImage, ByRef dstFormatParams As String, ByRef dstMetadataParams As String) As VbMsgBoxResult
    
    Load dialog_ExportAnimatedWebP
    dialog_ExportAnimatedWebP.ShowDialog srcImage
    
    PromptExportAnimatedWebP = dialog_ExportAnimatedWebP.GetDialogResult
    dstFormatParams = dialog_ExportAnimatedWebP.GetFormatParams
    dstMetadataParams = dialog_ExportAnimatedWebP.GetMetadataParams
    
    Unload dialog_ExportAnimatedWebP
    Set dialog_ExportAnimatedWebP = Nothing
    
End Function

Public Function PromptExportLUT(ByRef srcImage As pdImage, ByRef dstFormatParams As String) As VbMsgBoxResult
    
    Load dialog_ExportLUT
    dialog_ExportLUT.ShowDialog srcImage
    
    PromptExportLUT = dialog_ExportLUT.GetDialogResult
    dstFormatParams = dialog_ExportLUT.GetFormatParams
    
    Unload dialog_ExportLUT
    Set dialog_ExportLUT = Nothing
    
End Function

'Note: this function also works with WMF files
Public Function PromptImportEMF(ByVal hGdipImage As Long, ByVal origWidth As Long, ByVal origHeight As Long, ByRef dstWidth As Long, ByRef dstHeight As Long, ByRef dstDPI As Long) As VbMsgBoxResult
    
    Load dialog_ImportSVG
    dialog_ImportSVG.ShowDialog 0&, origWidth, origHeight, hGdipImage
    
    PromptImportEMF = dialog_ImportSVG.GetDialogResult
    dstWidth = dialog_ImportSVG.GetUserWidth
    dstHeight = dialog_ImportSVG.GetUserHeight
    dstDPI = dialog_ImportSVG.GetUserDPI
    
    Unload dialog_ImportSVG
    Set dialog_ImportSVG = Nothing
    
End Function

Public Function PromptImportSVG(ByVal hResvgTree As Long, ByVal origWidth As Long, ByVal origHeight As Long, ByRef dstWidth As Long, ByRef dstHeight As Long, ByRef dstDPI As Long) As VbMsgBoxResult
    
    Load dialog_ImportSVG
    dialog_ImportSVG.ShowDialog hResvgTree, origWidth, origHeight, 0&
    
    PromptImportSVG = dialog_ImportSVG.GetDialogResult
    dstWidth = dialog_ImportSVG.GetUserWidth
    dstHeight = dialog_ImportSVG.GetUserHeight
    dstDPI = dialog_ImportSVG.GetUserDPI
    
    Unload dialog_ImportSVG
    Set dialog_ImportSVG = Nothing
    
End Function

Public Function PromptImportPDF(ByRef cPDF As pdPDF, ByRef dstSettings As String) As VbMsgBoxResult

    Load dialog_ImportPDF
    dialog_ImportPDF.ShowDialog cPDF
    
    PromptImportPDF = dialog_ImportPDF.GetDialogResult()
    dstSettings = dialog_ImportPDF.GetDialogParamString()
    
    Unload dialog_ImportPDF
    Set dialog_ImportPDF = Nothing
    
End Function

Public Sub ShowClipboardDialog(ByVal clipMode As PD_ClipboardOp)

    Load FormClipboard
    FormClipboard.ShowClipboardDialog clipMode
    
    If (FormClipboard.DialogResult = vbOK) Then
        If (clipMode = co_Cut) Then
            Processor.Process "Cut special", False, FormClipboard.GetParamString(), UNDO_Image
        ElseIf (clipMode = co_Copy) Then
            Processor.Process "Copy special", False, FormClipboard.GetParamString(), UNDO_Nothing
        Else
            Processor.Process "Paste special", False, FormClipboard.GetParamString(), UNDO_Image_VectorSafe
        End If
    End If
    
    Unload FormClipboard
    Set FormClipboard = Nothing
    
End Sub
