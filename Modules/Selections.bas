Attribute VB_Name = "Selections"
'***************************************************************************
'Selection Interface
'Copyright 2013-2021 by Tanner Helland
'Created: 21/June/13
'Last updated: 15/February/21
'Last update: large selection tool overhaul to support multi-selection behavior
'
'Selection tools have existed in PhotoDemon for awhile, but this module is the first to support Process varieties of
' selection operations - e.g. internal actions like "Process "Create Selection"".  Selection commands must be passed
' through the Process module so they can be recorded as macros, and as part of the program's Undo/Redo chain.  This
' module provides all selection-related functions that the Process module can call.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_SelectionCombine
    pdsm_Replace = 0
    pdsm_Add = 1
    pdsm_Subtract = 2
    pdsm_Intersect = 3
End Enum

#If False Then
    Private Const pdsm_Replace = 0, pdsm_Add = 1, pdsm_Subtract = 2, pdsm_Intersect = 3
#End If

Public Enum PD_SelectionDialog
    pdsd_Grow = 0
    pdsd_Shrink = 1
    pdsd_Border = 2
    pdsd_Feather = 3
    pdsd_Sharpen = 4
End Enum

#If False Then
    Private Const pdsd_Grow = 0, pdsd_Shrink = 1, pdsd_Border = 2, pdsd_Feather = 3, pdsd_Sharpen = 4
#End If

Public Enum PD_SelectionRenderSetting
    pdsr_RenderMode = 0
    pdsr_HighlightColor = 1
    pdsr_HighlightOpacity = 2
    pdsr_LightboxColor = 3
    pdsr_LightboxOpacity = 4
End Enum

#If False Then
    Private Const pdsr_RenderMode = 0, pdsr_HighlightColor = 1, pdsr_HighlightOpacity = 2, pdsr_LightboxColor = 3, pdsr_LightboxOpacity = 4
#End If

'This module caches a number of UI-related selection details.  We cache these here because these
' are tied to program preferences (and not to specific selection instances).
Private m_CurSelectionMode As PD_SelectionRender, m_SelHighlightColor As Long, m_SelHighlightOpacity As Single
Private m_SelLightboxColor As Long, m_SelLightboxOpacity As Single

'A double-click event can be used to close the current polygon selection.  Unfortunately, this can
' have the (funny?) side-effect of removing the active selection, because the first click of the
' double-click causes a point to be created, but the second click causes that point to be removed
' and instead the polygon gets closed.  HOWEVER, on the subsequent _MouseUp, the click detector
' notices the _MouseUp potentially occurring *not* over the selection, and it erases the current
' selection accordingly.
'
'To avoid this debacle, we set a flag on the double-click event, and free it on the subsequent
' _MouseUp.
Private m_DblClickOccurred As Boolean

'Present a selection-related dialog box (grow, shrink, feather, etc).  This function will return a msgBoxResult value so
' the calling function knows how to proceed, and if the user successfully selected a value, it will be stored in the
' returnValue variable.
Public Function DisplaySelectionDialog(ByVal typeOfDialog As PD_SelectionDialog, ByRef ReturnValue As Double) As VbMsgBoxResult

    Load FormSelectionDialogs
    FormSelectionDialogs.ShowDialog typeOfDialog
    
    DisplaySelectionDialog = FormSelectionDialogs.DialogResult
    ReturnValue = FormSelectionDialogs.paramValue
    
    Unload FormSelectionDialogs
    Set FormSelectionDialogs = Nothing

End Function

'Create a new selection using the settings stored in a pdSerialize-compatible string
Public Sub CreateNewSelection(ByRef paramString As String)
    
    'Use the passed parameter string to initialize the selection
    PDImages.GetActiveImage.MainSelection.InitFromXML paramString
    PDImages.GetActiveImage.MainSelection.LockIn
    PDImages.GetActiveImage.SetSelectionActive True
    
    'For lasso selections, mark the lasso as closed if the selection is being created anew
    If (PDImages.GetActiveImage.MainSelection.GetSelectionShape() = ss_Lasso) Then PDImages.GetActiveImage.MainSelection.SetLassoClosedState True
    
    'Synchronize all user-facing controls to match
    Selections.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
End Sub

'Remove the current selection
Public Sub RemoveCurrentSelection(Optional ByVal updateUIToo As Boolean = True)
    
    'Release the selection object and mark it as inactive
    PDImages.GetActiveImage.MainSelection.LockRelease
    PDImages.GetActiveImage.SetSelectionActive False
    
    'Reset any internal selection state trackers
    PDImages.GetActiveImage.MainSelection.EraseCustomTrackers
    
    'Free as many unneeded caches as we can
    PDImages.GetActiveImage.MainSelection.FreeNonEssentialResources
    
    'Synchronize all user-facing controls to match
    If updateUIToo Then Selections.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
End Sub

'"Select all"
Public Sub SelectWholeImage()
    
    'Unselect any existing selection
    PDImages.GetActiveImage.MainSelection.LockRelease
    PDImages.GetActiveImage.SetSelectionActive False
    
    'Create a new selection at the size of the image
    PDImages.GetActiveImage.MainSelection.SelectAll
    
    'Lock in this selection
    PDImages.GetActiveImage.MainSelection.LockIn
    PDImages.GetActiveImage.SetSelectionActive True
    
    'Synchronize all user-facing controls to match
    SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
End Sub

'Load a previously saved selection.  Note that this function also handles creation and display of the relevant common dialog.
Public Sub LoadSelectionFromFile(ByVal displayDialog As Boolean, Optional ByVal loadSettings As String = vbNullString)

    If displayDialog Then
    
        'Disable user input until the dialog closes
        Interface.DisableUserInput
    
        'Simple open dialog
        Dim openDialog As pdOpenSaveDialog
        Set openDialog = New pdOpenSaveDialog
        
        Dim sFile As String
        
        Dim cdFilter As String
        cdFilter = g_Language.TranslateMessage("PhotoDemon selection") & " (.pds)|*.pds|"
        cdFilter = cdFilter & g_Language.TranslateMessage("All files") & "|*.*"
        
        Dim cdTitle As String
        cdTitle = g_Language.TranslateMessage("Load a previously saved selection")
                
        If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, UserPrefs.GetSelectionPath, cdTitle, , GetModalOwner().hWnd) Then
            
            'Use a temporary selection object to validate the requested selection file
            Dim tmpSelection As pdSelection
            Set tmpSelection = New pdSelection
            tmpSelection.SetParentReference PDImages.GetActiveImage()
            
            If tmpSelection.ReadSelectionFromFile(sFile, True) Then
                
                'Save the new directory as the default path for future usage
                UserPrefs.SetSelectionPath sFile
                
                'Call this function again, but with displayDialog set to FALSE and the path of the requested selection file
                Process "Load selection", False, BuildParamList("selectionpath", sFile), UNDO_Selection
                    
            Else
                PDMsgBox "An error occurred while attempting to load %1.  Please verify that the file is a valid PhotoDemon selection file.", vbOKOnly Or vbExclamation, "Error", sFile
            End If
            
            'Release the temporary selection object
            tmpSelection.SetParentReference Nothing
            Set tmpSelection = Nothing
            
        End If
        
        'Re-enable user input
        Interface.EnableUserInput
        
    Else
        
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString loadSettings
        
        Message "Loading selection..."
        PDImages.GetActiveImage.MainSelection.ReadSelectionFromFile cParams.GetString("selectionpath")
        PDImages.GetActiveImage.SetSelectionActive True
        
        'Synchronize all user-facing controls to match
        SyncTextToCurrentSelection PDImages.GetActiveImageID()
                
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        Message "Selection loaded successfully"
    
    End If
        
End Sub

'Save the current selection to file.  Note that this function also handles creation and display of the relevant common dialog.
Public Sub SaveSelectionToFile()

    'Simple save dialog
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    Dim sFile As String
    
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("PhotoDemon selection") & " (.pds)|*.pds"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save the current selection")
        
    If saveDialog.GetSaveFileName(sFile, , True, cdFilter, 1, UserPrefs.GetSelectionPath, cdTitle, ".pds", GetModalOwner().hWnd) Then
        
        'Save the new directory as the default path for future usage
        UserPrefs.SetSelectionPath sFile
        
        'Write out the selection file
        Dim cmpLevel As Long
        cmpLevel = Compression.GetDefaultCompressionLevel(cf_Zstd)
        If PDImages.GetActiveImage.MainSelection.WriteSelectionToFile(sFile, cf_Zstd, cmpLevel, cf_Zstd, cmpLevel) Then
            Message "Selection saved."
        Else
            Message "Unknown error occurred.  Selection was not saved.  Please try again."
        End If
        
    End If
        
End Sub

'Export the currently selected area as an image.  This is provided as a convenience to the user, so that they do not have to crop
' or copy-paste the selected area in order to save it.  The selected area is also checked for bit-depth; 24bpp is recommended as
' JPEG, while 32bpp is recommended as PNG (but the user can select any supported PD save format from the common dialog).
Public Function ExportSelectedAreaAsImage() As Boolean
    
    'If a selection is not active, it should be impossible to select this menu item.  Just in case, check for that state and exit if necessary.
    If (Not PDImages.GetActiveImage.IsSelectionActive) Then
        Message "This action requires an active selection.  Please create a selection before continuing."
        ExportSelectedAreaAsImage = False
        Exit Function
    End If
    
    'Prepare a temporary pdImage object to house the current selection mask
    Dim tmpImage As pdImage
    Set tmpImage = New pdImage
    
    'Copy the current selection DIB into a temporary DIB.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    PDImages.GetActiveImage.RetrieveProcessedSelection tmpDIB, True, True
    
    'In the temporary pdImage object, create a blank layer; this will receive the processed DIB
    Dim newLayerID As Long
    newLayerID = tmpImage.CreateBlankLayer
    tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, , tmpDIB
    tmpImage.UpdateSize
    
    'Give the selection a basic filename
    tmpImage.ImgStorage.AddEntry "OriginalFileName", "PhotoDemon selection"
    
    'Get the last "save image" path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    
    'By default, recommend JPEG for 24bpp selections, and PNG for 32bpp selections
    Dim saveFormat As Long
    If DIBs.IsDIBAlphaBinary(tmpDIB, False) Then
        saveFormat = ImageFormats.GetIndexOfOutputPDIF(PDIF_JPEG) + 1
    Else
        saveFormat = ImageFormats.GetIndexOfOutputPDIF(PDIF_PNG) + 1
    End If
    
    'Now it's time to prepare a standard Save Image common dialog
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    'Provide a string to the common dialog; it will fill this with the user's chosen path + filename
    Dim sFile As String
    sFile = tempPathString & IncrementFilename(tempPathString, tmpImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString), ImageFormats.GetOutputFormatExtension(saveFormat - 1))
    
    'Present a common dialog to the user
    If saveDialog.GetSaveFileName(sFile, , True, ImageFormats.GetCommonDialogOutputFormats, saveFormat, tempPathString, g_Language.TranslateMessage("Export selection as image"), ImageFormats.GetCommonDialogDefaultExtensions, FormMain.hWnd) Then
        
        'Store the selected file format to the image object
        tmpImage.SetCurrentFileFormat ImageFormats.GetOutputPDIF(saveFormat - 1)
        
        'Transfer control to the core SaveImage routine, which will handle color depth analysis and actual saving
        ExportSelectedAreaAsImage = PhotoDemon_SaveImage(tmpImage, sFile, True)
        
    Else
        ExportSelectedAreaAsImage = False
    End If
        
    'Release our temporary image
    Set tmpDIB = Nothing
    Set tmpImage = Nothing
    
End Function

'Export the current selection mask as an image.  PNG is recommended by default, but the user can choose from any of PD's available formats.
Public Function ExportSelectionMaskAsImage() As Boolean
    
    'If a selection is not active, it should be impossible to select this menu item.  Just in case, check for that state and exit if necessary.
    If Not PDImages.GetActiveImage.IsSelectionActive Then
        Message "This action requires an active selection.  Please create a selection before continuing."
        ExportSelectionMaskAsImage = False
        Exit Function
    End If
    
    'Prepare a temporary pdImage object to house the current selection mask
    Dim tmpImage As pdImage
    Set tmpImage = New pdImage
    
    'Create a temporary DIB, then retrieve the current selection into it
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB PDImages.GetActiveImage.MainSelection.GetMaskDIB
    
    'Selections use a "white = selected, transparent = unselected" strategy.  Composite against a black background now
    ' (but leave the DIB in 32-bpp format)
    tmpDIB.CompositeBackgroundColor 0, 0, 0
    
    'In a temporary pdImage object, create a blank layer; this will receive the processed DIB
    Dim newLayerID As Long
    newLayerID = tmpImage.CreateBlankLayer
    tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, , tmpDIB
    tmpImage.UpdateSize
    
    'Give the selection a basic filename
    tmpImage.ImgStorage.AddEntry "OriginalFileName", g_Language.TranslateMessage("PhotoDemon selection")
        
    'Get the last "save image" path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    
    'By default, recommend PNG as the save format
    Dim saveFormat As Long
    saveFormat = ImageFormats.GetIndexOfOutputPDIF(PDIF_PNG) + 1
    
    'Provide a string to the common dialog; it will fill this with the user's chosen path + filename
    Dim sFile As String
    sFile = tempPathString & IncrementFilename(tempPathString, tmpImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString), "png")
    
    'Now it's time to prepare a standard Save Image common dialog
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    'Present a common dialog to the user
    If saveDialog.GetSaveFileName(sFile, , True, ImageFormats.GetCommonDialogOutputFormats, saveFormat, tempPathString, g_Language.TranslateMessage("Export selection as image"), ImageFormats.GetCommonDialogDefaultExtensions, FormMain.hWnd) Then
                
        'Store the selected file format to the image object
        tmpImage.SetCurrentFileFormat ImageFormats.GetOutputPDIF(saveFormat - 1)
                                
        'Transfer control to the core SaveImage routine, which will handle color depth analysis and actual saving
        ExportSelectionMaskAsImage = PhotoDemon_SaveImage(tmpImage, sFile, True)
        
    Else
        ExportSelectionMaskAsImage = False
    End If
    
    'Release our temporary image
    Set tmpImage = Nothing

End Function

'Use this to populate the text boxes on the main form with the current selection values.
' (Note that this does not cause a screen refresh, by design.)
Public Sub SyncTextToCurrentSelection(ByVal srcImageID As Long)

    Dim i As Long
    
    'Only synchronize the text boxes if a selection is active
    Dim selectionIsActive As Boolean
    selectionIsActive = Selections.SelectionsAllowed(False)
    
    Dim selectionToolActive As Boolean
    If selectionIsActive Then
        If PDImages.IsImageActive(srcImageID) Then selectionToolActive = Tools.IsSelectionToolActive()
    End If
    
    'See if a selection exists
    If selectionIsActive And selectionToolActive Then
        
        PDImages.GetImageByID(srcImageID).MainSelection.SuspendAutoRefresh True
    
        'Selection coordinate toolboxes appear on three different selection subpanels: rect, ellipse, and line.
        ' To access their indicies properly, we must calculate an offset.
        Dim subpanelOffset As Long, subpanelCtlOffset As Long
        subpanelOffset = Selections.GetSelectionSubPanelFromSelectionShape(PDImages.GetImageByID(srcImageID))
        subpanelCtlOffset = subpanelOffset * 2
        
        'Additional syncing is done if the selection is transformable; if it is not transformable, clear and lock the location text boxes
        If PDImages.GetImageByID(srcImageID).MainSelection.IsTransformable Then
            
            Dim tmpRectF As RectF, tmpRectFRB As RectF_RB
            
            'Different types of selections will display size and position differently
            Select Case PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape()
                
                'Rectangular and elliptical selections display left, top, width, height, and aspect ratio (in the form X:Y)
                Case ss_Rectangle, ss_Circle
                    
                    Dim sizeIndex As Long
                    sizeIndex = toolpanel_Selections.cboSize(subpanelOffset).ListIndex
                    
                    'Coordinates are allowed to be <= 0, but size and aspect ratio are not
                    Dim allowedMin As Long
                    If (sizeIndex = 0) Then allowedMin = -32000 Else allowedMin = 1
                    If (toolpanel_Selections.tudSel(subpanelCtlOffset + 0).Min <> allowedMin) Then
                        toolpanel_Selections.tudSel(subpanelCtlOffset + 0).Min = allowedMin
                        toolpanel_Selections.tudSel(subpanelCtlOffset + 1).Min = allowedMin
                    End If
                    
                    tmpRectF = PDImages.GetImageByID(srcImageID).MainSelection.GetCornersLockedRect()
                    If (sizeIndex = 0) Then
                        toolpanel_Selections.tudSel(subpanelCtlOffset + 0).Value = tmpRectF.Left
                        toolpanel_Selections.tudSel(subpanelCtlOffset + 1).Value = tmpRectF.Top
                    ElseIf (sizeIndex = 1) Then
                        toolpanel_Selections.tudSel(subpanelCtlOffset + 0).Value = tmpRectF.Width
                        toolpanel_Selections.tudSel(subpanelCtlOffset + 1).Value = tmpRectF.Height
                    ElseIf (sizeIndex = 2) Then
                        
                        'Failsafe DBZ check
                        If (tmpRectF.Height > 0) Then
                        
                            Dim fracNumerator As Long, fracDenominator As Long
                            PDMath.ConvertToFraction tmpRectF.Width / tmpRectF.Height, fracNumerator, fracDenominator, 0.005
                            
                            'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
                            If (fracDenominator = 5) Then
                                fracNumerator = fracNumerator * 2
                                fracDenominator = fracDenominator * 2
                            End If
                            
                            toolpanel_Selections.tudSel(subpanelCtlOffset + 0).Value = fracNumerator
                            toolpanel_Selections.tudSel(subpanelCtlOffset + 1).Value = fracDenominator
                            
                        End If
                        
                    End If
                    
                    '"Lock" button visibility is a little complicated - basically, we only want to make it visible
                    ' for width, height, and aspect ratio options.
                    If (sizeIndex = 0) Then
                        toolpanel_Selections.cmdLock(subpanelOffset * 2).Visible = False
                        toolpanel_Selections.cmdLock(subpanelOffset * 2 + 1).Visible = False
                        toolpanel_Selections.lblColon(subpanelOffset).Visible = False
                    ElseIf (sizeIndex = 1) Then
                        toolpanel_Selections.cmdLock(subpanelOffset * 2).Visible = True
                        toolpanel_Selections.cmdLock(subpanelOffset * 2 + 1).Visible = True
                        toolpanel_Selections.lblColon(subpanelOffset).Visible = False
                    Else
                        toolpanel_Selections.cmdLock(subpanelOffset * 2).Visible = False
                        toolpanel_Selections.cmdLock(subpanelOffset * 2 + 1).Visible = True
                        toolpanel_Selections.lblColon(subpanelOffset).Visible = True
                    End If
                    
                    'Also make sure the "lock" icon matches the current lock state
                    If (sizeIndex = 1) Then
                        toolpanel_Selections.cmdLock(subpanelCtlOffset).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetPropertyLockedState(pdsl_Width)
                        toolpanel_Selections.cmdLock(subpanelCtlOffset + 1).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetPropertyLockedState(pdsl_Height)
                    ElseIf (sizeIndex = 2) Then
                        toolpanel_Selections.cmdLock(subpanelCtlOffset + 1).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetPropertyLockedState(pdsl_AspectRatio)
                    End If
                    
            End Select
            
        Else
        
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                If (toolpanel_Selections.tudSel(i).Value <> 0) Then toolpanel_Selections.tudSel(i).Value = 0
            Next i
            
        End If
        
        'Next, sync all non-coordinate information
        If (PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape <> ss_Raster) And (PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape <> ss_Wand) Then
            toolpanel_Selections.cboSelArea(Selections.GetSelectionSubPanelFromSelectionShape(PDImages.GetImageByID(srcImageID))).ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_Area)
            toolpanel_Selections.sltSelectionBorder(Selections.GetSelectionSubPanelFromSelectionShape(PDImages.GetImageByID(srcImageID))).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_BorderWidth)
        End If
        
        If toolpanel_Selections.cboSelSmoothing.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_Smoothing) Then toolpanel_Selections.cboSelSmoothing.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_Smoothing)
        If toolpanel_Selections.sltSelectionFeathering.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_FeatheringRadius) Then toolpanel_Selections.sltSelectionFeathering.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_FeatheringRadius)
        
        'Finally, sync any shape-specific information
        Select Case PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape
        
            Case ss_Rectangle
                If (toolpanel_Selections.sltCornerRounding.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_RoundedCornerRadius)) Then toolpanel_Selections.sltCornerRounding.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_RoundedCornerRadius)
            
            Case ss_Circle
            
            Case ss_Lasso
                If toolpanel_Selections.sltSmoothStroke.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_SmoothStroke) Then toolpanel_Selections.sltSmoothStroke.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_SmoothStroke)
                
            Case ss_Polygon
                If toolpanel_Selections.sltPolygonCurvature.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_PolygonCurvature) Then toolpanel_Selections.sltPolygonCurvature.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_PolygonCurvature)
                
            Case ss_Wand
                If toolpanel_Selections.btsWandArea.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSearchMode) Then toolpanel_Selections.btsWandArea.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSearchMode)
                If toolpanel_Selections.btsWandMerge.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSampleMerged) Then toolpanel_Selections.btsWandMerge.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSampleMerged)
                If toolpanel_Selections.sltWandTolerance.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_WandTolerance) Then toolpanel_Selections.sltWandTolerance.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_WandTolerance)
                If toolpanel_Selections.cboWandCompare.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandCompareMethod) Then toolpanel_Selections.cboWandCompare.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandCompareMethod)
        
        End Select
        
        PDImages.GetImageByID(srcImageID).MainSelection.SuspendAutoRefresh False
    
    'A selection is *not* active; disable various selection-related UI options
    Else
        
        'If a selection exists, we need to leave available menu commands like "remove selection", etc.
        Interface.SetUIGroupState PDUI_Selections, selectionIsActive
        
        'Transformable settings do *not* need to be available
        Interface.SetUIGroupState PDUI_SelectionTransforms, False
        
        'This branch is only followed if a selection is *not* active but a selection tool *is* active, in which case
        ' we need to disable some commands on the selection toolbar.
        If Tools.IsSelectionToolActive Then
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                If (toolpanel_Selections.tudSel(i).Value <> 0) Then toolpanel_Selections.tudSel(i).Value = 0
            Next i
            For i = 0 To toolpanel_Selections.cmdLock.Count - 1
                toolpanel_Selections.cmdLock(i).Visible = False
            Next i
            For i = 0 To toolpanel_Selections.lblColon.Count - 1
                toolpanel_Selections.lblColon(i).Visible = False
            Next i
        End If
        
    End If
    
End Sub

'Given an (x, y) pair in IMAGE coordinate space (not screen or canvas space), return a constant if the point is a valid
' "point of interest" for the active selection.  Standard UI mouse distances are allowed (meaning zoom is factored into the
' algorithm).
'
'The result of this function is typically passed to something like pdSelection.SetActiveSelectionPOI(), which will cache
' the point of interest and use it to interpret subsequent mouse events (e.g. click-dragging a selection to a new position).
'
'Note that only certain POIs are hard-coded.  Some selections (e.g. polygons) can return other values outside the enum,
' typically indices into an internal selection point array.
'
'This sub will return a constant correlating to the nearest selection point.  See the relevant enum for details.
Public Function IsCoordSelectionPOI(ByVal imgX As Double, ByVal imgY As Double, ByRef srcImage As pdImage) As PD_PointOfInterest
    
    'If the current selection is...
    ' 1) raster-type, or...
    ' 2) inactive...
    '...disallow POIs entirely.  (These types of selections do not support on-canvas interactions.)
    If (srcImage.MainSelection.GetSelectionShape = ss_Raster) Or (Not srcImage.IsSelectionActive) Then
        IsCoordSelectionPOI = poi_Undefined
        Exit Function
    End If
    
    'Similarly, POIs are only enabled if the current selection tool matches the current selection shape
    If (g_CurrentTool <> Selections.GetRelevantToolFromSelectShape()) Then
        IsCoordSelectionPOI = poi_Undefined
        Exit Function
    End If
    
    'We're now going to compare the passed coordinate against a hard-coded list of "points of interest."  These POIs
    ' differ by selection type, as different selections allow for different levels of interaction.  (For example, a polygon
    ' selection behaves differently when a point is dragged, vs a rectangular selection.)
    
    'Regardless of selection type, start by establishing boundaries for the current selection.
    'Calculate points of interest for the current selection.  Individual selection types define what is considered a POI,
    ' but in most cases, corners or interior clicks tend to allow some kind of user interaction.
    Dim tmpRectF As RectF
    If (srcImage.MainSelection.GetSelectionShape = ss_Rectangle) Or (srcImage.MainSelection.GetSelectionShape = ss_Circle) Then
        tmpRectF = srcImage.MainSelection.GetCornersLockedRect()
    Else
        tmpRectF = srcImage.MainSelection.GetBoundaryRect()
    End If
    
    'Adjust the mouseAccuracy value based on the current zoom value
    Dim mouseAccuracy As Double
    mouseAccuracy = Drawing.ConvertCanvasSizeToImageSize(Interface.GetStandardInteractionDistance(), srcImage)
        
    'Find the smallest distance for this mouse position
    Dim minDistance As Double
    minDistance = mouseAccuracy
    
    Dim closestPoint As Long
    closestPoint = poi_Undefined
    
    'Some selection types (lasso, polygon) must use a more complicated region for hit-testing.  GDI+ will be used for this.
    Dim complexRegion As pd2DRegion
    
    'Other selection types will use a generic list of points (like the corners of the current selection)
    Dim poiListFloat() As PointFloat
    
    'If we made it here, this mouse location is worth evaluating.  How we evaluate it depends on the shape of the current selection.
    Select Case srcImage.MainSelection.GetSelectionShape
    
        'Rectangular and elliptical selections have identical POIs: the corners, edges, and interior of the selection
        Case ss_Rectangle, ss_Circle
    
            'Corners get preference, so check them first.
            ReDim poiListFloat(0 To 3) As PointFloat
            
            With tmpRectF
                poiListFloat(0).x = .Left
                poiListFloat(0).y = .Top
                poiListFloat(1).x = .Left + .Width
                poiListFloat(1).y = .Top
                poiListFloat(2).x = .Left + .Width
                poiListFloat(2).y = .Top + .Height
                poiListFloat(3).x = .Left
                poiListFloat(3).y = .Top + .Height
            End With
            
            'Used the generalized point comparison function to see if one of the points matches
            closestPoint = FindClosestPointInFloatArray(imgX, imgY, minDistance, poiListFloat)
            
            'Did one of the corner points match?  If so, map it to a valid constant and return.
            If (closestPoint <> poi_Undefined) Then
                
                If (closestPoint = 0) Then
                    IsCoordSelectionPOI = poi_CornerNW
                ElseIf (closestPoint = 1) Then
                    IsCoordSelectionPOI = poi_CornerNE
                ElseIf (closestPoint = 2) Then
                    IsCoordSelectionPOI = poi_CornerSE
                ElseIf (closestPoint = 3) Then
                    IsCoordSelectionPOI = poi_CornerSW
                End If
                
            Else
        
                'If we're at this line of code, a closest corner was not found.  Check edges next.
                ' (Unfortunately, we don't yet have a generalized function for edge checking, so this must be done manually.)
                '
                'Note that edge checks are a little weird currently, because we check one-dimensional distance between each
                ' side, and if that's a hit, we see if the point also lies between the bounds in the *other* direction.
                ' This allows the user to use the entire selection side to perform a stretch.
                Dim nDist As Double, eDist As Double, sDist As Double, wDist As Double
                
                With tmpRectF
                    nDist = DistanceOneDimension(imgY, .Top)
                    eDist = DistanceOneDimension(imgX, .Left + .Width)
                    sDist = DistanceOneDimension(imgY, .Top + .Height)
                    wDist = DistanceOneDimension(imgX, .Left)
                
                    If (nDist <= minDistance) Then
                        If (imgX > (.Left - minDistance)) And (imgX < (.Left + .Width + minDistance)) Then
                            minDistance = nDist
                            closestPoint = poi_EdgeN
                        End If
                    End If
                    
                    If (eDist <= minDistance) Then
                        If (imgY > (.Top - minDistance)) And (imgY < (.Top + .Height + minDistance)) Then
                            minDistance = eDist
                            closestPoint = poi_EdgeE
                        End If
                    End If
                    
                    If (sDist <= minDistance) Then
                        If (imgX > (.Left - minDistance)) And (imgX < (.Left + .Width + minDistance)) Then
                            minDistance = sDist
                            closestPoint = poi_EdgeS
                        End If
                    End If
                    
                    If (wDist <= minDistance) Then
                        If (imgY > (.Top - minDistance)) And (imgY < (.Top + .Height + minDistance)) Then
                            minDistance = wDist
                            closestPoint = poi_EdgeW
                        End If
                    End If
                
                End With
                
                'Was a close point found? If yes, then return that value.
                If (closestPoint <> poi_Undefined) Then
                    IsCoordSelectionPOI = closestPoint
                Else
            
                    'If we're at this line of code, a closest edge was not found. Perform one final check to ensure that the mouse is within the
                    ' image's boundaries, and if it is, return the "move selection" ID, then exit.
                    If PDMath.IsPointInRectF(imgX, imgY, tmpRectF) Then
                        IsCoordSelectionPOI = poi_Interior
                    Else
                        IsCoordSelectionPOI = poi_Undefined
                    End If
                    
                End If
                
            End If
            
        Case ss_Polygon
        
            'First, we want to check all polygon points for a hit.
            PDImages.GetActiveImage.MainSelection.GetPolygonPoints poiListFloat()
            
            'Used the generalized point comparison function to see if one of the points matches
            closestPoint = FindClosestPointInFloatArray(imgX, imgY, minDistance, poiListFloat)
            
            'Was a close point found? If yes, then return that value
            If (closestPoint <> poi_Undefined) Then
                IsCoordSelectionPOI = closestPoint
                
            'If no polygon point was a hit, our final check is to see if the mouse lies within the polygon itself.  This will trigger
            ' a move transformation.
            Else
                
                'Use a region object for hit-detection
                Set complexRegion = PDImages.GetActiveImage.MainSelection.GetSelectionAsRegion()
                If (Not complexRegion Is Nothing) Then
                    If complexRegion.IsPointInRegion(imgX, imgY) Then IsCoordSelectionPOI = poi_Interior Else IsCoordSelectionPOI = poi_Undefined
                Else
                    IsCoordSelectionPOI = poi_Undefined
                End If
                
            End If
        
        Case ss_Lasso
        
            'Use a region object for hit-detection
            Set complexRegion = PDImages.GetActiveImage.MainSelection.GetSelectionAsRegion()
            If (Not complexRegion Is Nothing) Then
                If complexRegion.IsPointInRegion(imgX, imgY) Then IsCoordSelectionPOI = poi_Interior Else IsCoordSelectionPOI = poi_Undefined
            Else
                IsCoordSelectionPOI = poi_Undefined
            End If
                
        Case ss_Wand
            
            'Wand selections do actually support a single point of interest - the wand's "clicked" location
            srcImage.MainSelection.GetCurrentPOIList poiListFloat
            
            'Used the generalized point comparison function to see if one of the points matches
            IsCoordSelectionPOI = FindClosestPointInFloatArray(imgX, imgY, minDistance, poiListFloat)
            
        Case Else
            IsCoordSelectionPOI = poi_Undefined
            Exit Function
            
    End Select

End Function

'Invert the current selection.  Note that this will make a transformable selection non-transformable - to maintain transformability, use
' the "exterior"/"interior" options on the main form.
' TODO: swap exterior/interior automatically, if a valid option
Public Sub InvertCurrentSelection()

    'Unselect any existing selection
    PDImages.GetActiveImage.MainSelection.LockRelease
    PDImages.GetActiveImage.SetSelectionActive False
        
    Message "Inverting..."
    
    'Point a standard 2D byte array at the selection mask
    Dim x As Long, y As Long
    Dim selMaskData() As Long, selMaskSA As SafeArray1D
    
    Dim maskWidth As Long, maskHeight As Long
    maskWidth = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth - 1
    maskHeight = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    SetProgBarMax maskHeight
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'After all that work, the Invert code itself is very small and unexciting!
    For y = 0 To maskHeight
        PDImages.GetActiveImage.MainSelection.GetMaskDIB.WrapLongArrayAroundScanline selMaskData, selMaskSA, y
    For x = 0 To maskWidth
        selMaskData(x) = Not selMaskData(x)
    Next x
        If (y And progBarCheck) = 0 Then SetProgBarVal y
    Next y
    
    PDImages.GetActiveImage.MainSelection.GetMaskDIB.UnwrapLongArrayFromDIB selMaskData
    
    'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
    ' being non-transformable)
    PDImages.GetActiveImage.MainSelection.SetSelectionShape ss_Raster
    PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
    PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
    
    SetProgBarVal 0
    ReleaseProgressBar
    Message "Selection inversion complete."
    
    'Lock in this selection
    PDImages.GetActiveImage.MainSelection.LockIn
    PDImages.GetActiveImage.SetSelectionActive True
        
    'Draw the new selection to the screen
    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Feather the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub FeatherCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal featherRadius As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retRadius As Double
        If DisplaySelectionDialog(pdsd_Feather, retRadius) = vbOK Then
            Process "Feather selection", False, BuildParamList("filtervalue", retRadius), UNDO_Selection
        End If
        
    Else
    
        Message "Feathering selection..."
    
        'Unselect any existing selection
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        
        'Retrieve just the alpha channel of the current selection
        Dim tmpArray() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetMaskDIB, tmpArray
        
        'Blur that temporary array
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight
        Filters_ByteArray.HorizontalBlur_ByteArray tmpArray, arrWidth, arrHeight, featherRadius, featherRadius
        Filters_ByteArray.VerticalBlur_ByteArray tmpArray, arrWidth, arrHeight, featherRadius, featherRadius
        
        'Reconstruct the DIB from the transparency table
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetMaskDIB, tmpArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Feathering complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If

End Sub

'Sharpen (un-feather?) the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub SharpenCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal sharpenRadius As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retRadius As Double
        If (DisplaySelectionDialog(pdsd_Sharpen, retRadius) = vbOK) Then
            Process "Sharpen selection", False, BuildParamList("filtervalue", retRadius), UNDO_Selection
        End If
        
    Else
    
        Message "Sharpening selection..."
    
        'Unselect any existing selection
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
                
        'Retrieve just the alpha channel of the current selection, and clone it so that we have two copies
        Dim tmpArray() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetMaskDIB, tmpArray
        
        Dim tmpDstArray() As Byte
        ReDim tmpDstArray(0 To PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth - 1, PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight - 1) As Byte
        CopyMemoryStrict VarPtr(tmpDstArray(0, 0)), VarPtr(tmpArray(0, 0)), PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth * PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight
        
        'Blur the first temporary array
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight
        Filters_ByteArray.HorizontalBlur_ByteArray tmpArray, arrWidth, arrHeight, sharpenRadius, sharpenRadius
        Filters_ByteArray.VerticalBlur_ByteArray tmpArray, arrWidth, arrHeight, sharpenRadius, sharpenRadius
        
        'We're now going to perform an "unsharp mask" effect, but because we're using a single channel, it goes a bit faster
        Dim progBarCheck As Long
        SetProgBarMax PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight
        progBarCheck = ProgressBars.FindBestProgBarValue()
        
        'ScaleFactor is used to apply the unsharp mask.  Maximum strength can be any value, but PhotoDemon locks it at 10
        ' for selections (which are predictably feathered, using exact gaussian techniques).
        Dim scaleFactor As Double, invScaleFactor As Double
        scaleFactor = sharpenRadius
        invScaleFactor = 1# - scaleFactor
        
        Dim iWidth As Long, iHeight As Long
        iWidth = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth - 1
        iHeight = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight - 1
        
        Dim lOrig As Long, lBlur As Long, lDelta As Single, lFull As Single, lNew As Long
        Dim x As Long, y As Long
        
        Const ONE_DIV_255 As Double = 1# / 255#
        
        For y = 0 To iHeight
        For x = 0 To iWidth
            
            'Retrieve the original and blurred byte values
            lOrig = tmpDstArray(x, y)
            lBlur = tmpArray(x, y)
            
            'Calculate the delta between the two, which is then converted to a blend factor
            lDelta = Abs(lOrig - lBlur) * ONE_DIV_255
            
            'Calculate a "fully" sharpened value; we're going to manually feather between this value and the original,
            ' based on the delta between the two.
            lFull = (scaleFactor * lOrig) + (invScaleFactor * lBlur)
            
            'Feather to arrive at a final "unsharp" value
            lNew = (1# - lDelta) * lFull + (lDelta * lOrig)
            If (lNew < 0) Then
                lNew = 0
            ElseIf (lNew > 255) Then
                lNew = 255
            End If
            
            'Since we're doing a per-pixel loop, we can safely store the result back into the destination array
            tmpDstArray(x, y) = lNew
            
        Next x
            If (x And progBarCheck) = 0 Then SetProgBarVal y
        Next y
        
        'Reconstruct the DIB from the finished transparency table
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetMaskDIB, tmpDstArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Feathering complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If

End Sub

'Grow the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub GrowCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal growSize As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retSize As Double
        If DisplaySelectionDialog(pdsd_Grow, retSize) = vbOK Then
            Process "Grow selection", False, BuildParamList("filtervalue", retSize), UNDO_Selection
        End If
        
    Else
    
        Message "Growing selection..."
    
        'Unselect any existing selection
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        
        'Use PD's built-in Median function to dilate the selected area
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight
        
        Dim tmpArray() As Byte
        ReDim tmpArray(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        
        Dim srcBytes() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetMaskDIB, srcBytes
        
        If Filters_ByteArray.Dilate_ByteArray(growSize, PDPRS_Circle, srcBytes, tmpArray, arrWidth, arrHeight) Then
            DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetMaskDIB, tmpArray
        End If
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Selection resize complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If
    
End Sub

'Shrink the current selection.  Note that this will make a transformable selection non-transformable.
Public Sub ShrinkCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal shrinkSize As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retSize As Double
        If DisplaySelectionDialog(pdsd_Shrink, retSize) = vbOK Then
            Process "Shrink selection", False, BuildParamList("filtervalue", retSize), UNDO_Selection
        End If
        
    Else
    
        Message "Shrinking selection..."
    
        'Unselect any existing selection
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        
        'Use PD's built-in Median function to dilate the selected area
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight
        
        Dim tmpArray() As Byte
        ReDim tmpArray(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        
        Dim srcBytes() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetMaskDIB, srcBytes
        
        Filters_ByteArray.Erode_ByteArray shrinkSize, PDPRS_Circle, srcBytes, tmpArray, arrWidth, arrHeight
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetMaskDIB, tmpArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Selection resize complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If
    
End Sub

'Convert the current selection to border-type.  Note that this will make a transformable selection non-transformable.
Public Sub BorderCurrentSelection(ByVal displayDialog As Boolean, Optional ByVal borderRadius As Double = 0#)

    'If a dialog has been requested, display one to the user.  Otherwise, proceed with the feathering.
    If displayDialog Then
        
        Dim retSize As Double
        If DisplaySelectionDialog(pdsd_Border, retSize) = vbOK Then
            Process "Border selection", False, BuildParamList("filtervalue", retSize), UNDO_Selection
        End If
        
    Else
    
        Message "Finding selection border..."
    
        'Unselect any existing selection
        PDImages.GetActiveImage.MainSelection.LockRelease
        PDImages.GetActiveImage.SetSelectionActive False
        
        'Bordering a selection requires two passes: a grow pass and a shrink pass.  The results of these two passes are then blended
        ' to create the final bordered selection.
        
        'First, extract selection data into a byte array so we can use optimized analysis functions
        Dim arrWidth As Long, arrHeight As Long
        arrWidth = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth
        arrHeight = PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBHeight
        
        Dim srcArray() As Byte
        DIBs.RetrieveTransparencyTable PDImages.GetActiveImage.MainSelection.GetMaskDIB, srcArray
        
        'Next, generate a shrink (erode) pass
        Dim shrinkBytes() As Byte
        ReDim shrinkBytes(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        Filters_ByteArray.Erode_ByteArray borderRadius, PDPRS_Circle, srcArray, shrinkBytes, arrWidth, arrHeight, False, PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth * 2
        
        'Generate a grow (dilate) pass
        Dim growBytes() As Byte
        ReDim growBytes(0 To arrWidth - 1, 0 To arrHeight - 1) As Byte
        Filters_ByteArray.Dilate_ByteArray borderRadius, PDPRS_Circle, srcArray, growBytes, arrWidth, arrHeight, False, PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth * 2, PDImages.GetActiveImage.MainSelection.GetMaskDIB.GetDIBWidth
        
        'Finally, XOR those results together: that's our border!
        Dim x As Long, y As Long
        For y = 0 To arrHeight - 1
        For x = 0 To arrWidth - 1
            srcArray(x, y) = shrinkBytes(x, y) Xor growBytes(x, y)
        Next x
        Next y
        
        'Reconstruct the target DIB from our final array
        DIBs.Construct32bppDIBFromByteMap PDImages.GetActiveImage.MainSelection.GetMaskDIB, srcArray
        
        'Ask the selection to find new boundaries.  This will also set all relevant parameters for the modified selection (such as
        ' being non-transformable)
        PDImages.GetActiveImage.MainSelection.NotifyRasterDataChanged
        PDImages.GetActiveImage.MainSelection.FindNewBoundsManually
                
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
                
        SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Selection resize complete."
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If
    
End Sub

'Erase the currently selected area (LAYER ONLY!).  Note that this will not modify the current selection in any way;
' only the layer's pixel contents will be affected.
Public Sub EraseSelectedArea(ByVal targetLayerIndex As Long)
    PDImages.GetActiveImage.EraseProcessedSelection targetLayerIndex
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

'The selection engine integrates closely with tool selection (as it needs to know what kind of selection is being
' created/edited at any given time).  This function is called whenever the selection engine needs to correlate the
' current tool with a selection shape.  This allows us to easily switch between a rectangle and circle selection,
' for example, without forcing the user to recreate the selection from scratch.
Public Function GetSelectionShapeFromCurrentTool() As PD_SelectionShape

    Select Case g_CurrentTool
    
        Case SELECT_RECT
            GetSelectionShapeFromCurrentTool = ss_Rectangle
            
        Case SELECT_CIRC
            GetSelectionShapeFromCurrentTool = ss_Circle
        
        Case SELECT_POLYGON
            GetSelectionShapeFromCurrentTool = ss_Polygon
            
        Case SELECT_LASSO
            GetSelectionShapeFromCurrentTool = ss_Lasso
            
        Case SELECT_WAND
            GetSelectionShapeFromCurrentTool = ss_Wand
            
        Case Else
            GetSelectionShapeFromCurrentTool = -1
    
    End Select
    
End Function

'The inverse of "getSelectionShapeFromCurrentTool", above
Public Function GetRelevantToolFromSelectShape() As PDTools

    If PDImages.IsImageActive() Then

        If (Not PDImages.GetActiveImage.MainSelection Is Nothing) Then

            Select Case PDImages.GetActiveImage.MainSelection.GetSelectionShape
            
                Case ss_Rectangle
                    GetRelevantToolFromSelectShape = SELECT_RECT
                    
                Case ss_Circle
                    GetRelevantToolFromSelectShape = SELECT_CIRC
                
                Case ss_Polygon
                    GetRelevantToolFromSelectShape = SELECT_POLYGON
                    
                Case ss_Lasso
                    GetRelevantToolFromSelectShape = SELECT_LASSO
                    
                Case ss_Wand
                    GetRelevantToolFromSelectShape = SELECT_WAND
                
                Case Else
                    GetRelevantToolFromSelectShape = -1
            
            End Select
            
        Else
            GetRelevantToolFromSelectShape = -1
        End If
            
    Else
        GetRelevantToolFromSelectShape = -1
    End If

End Function

'All selection tools share the same main panel on the options toolbox, but they have different subpanels that contain their
' specific parameters.  Use this function to correlate the two.
Public Function GetSelectionSubPanelFromCurrentTool() As Long

    Select Case g_CurrentTool
    
        Case SELECT_RECT
            GetSelectionSubPanelFromCurrentTool = 0
            
        Case SELECT_CIRC
            GetSelectionSubPanelFromCurrentTool = 1
        
        Case SELECT_POLYGON
            GetSelectionSubPanelFromCurrentTool = 2
            
        Case SELECT_LASSO
            GetSelectionSubPanelFromCurrentTool = 3
            
        Case SELECT_WAND
            GetSelectionSubPanelFromCurrentTool = 4
        
        Case Else
            GetSelectionSubPanelFromCurrentTool = -1
    
    End Select
    
End Function

Public Function GetSelectionSubPanelFromSelectionShape(ByRef srcImage As pdImage) As Long

    Select Case srcImage.MainSelection.GetSelectionShape
    
        Case ss_Rectangle
            GetSelectionSubPanelFromSelectionShape = 0
            
        Case ss_Circle
            GetSelectionSubPanelFromSelectionShape = 1
        
        Case ss_Polygon
            GetSelectionSubPanelFromSelectionShape = 2
            
        Case ss_Lasso
            GetSelectionSubPanelFromSelectionShape = 3
            
        Case ss_Wand
            GetSelectionSubPanelFromSelectionShape = 4
        
        Case Else
            GetSelectionSubPanelFromSelectionShape = -1
    
    End Select
    
End Function

'Selections can be initiated several different ways.  To cut down on duplicated code, all new selection instances are referred
' to this function.  Initial X/Y values are required.
Public Sub InitSelectionByPoint(ByVal x As Double, ByVal y As Double)
    
    'Reset any existing selection properties
    PDImages.GetActiveImage.MainSelection.EraseCustomTrackers
    
    'Activate the attached image's primary selection
    PDImages.GetActiveImage.SetSelectionActive True
    PDImages.GetActiveImage.MainSelection.LockRelease
    
    'Reflect all current selection tool settings to the active selection object
    Dim curShape As PD_SelectionShape
    curShape = Selections.GetSelectionShapeFromCurrentTool()
    
    With PDImages.GetActiveImage.MainSelection
        .SetSelectionShape curShape
        If (curShape <> ss_Wand) Then .SetSelectionProperty sp_Area, toolpanel_Selections.cboSelArea(Selections.GetSelectionSubPanelFromCurrentTool).ListIndex Else .SetSelectionProperty sp_Area, sa_Interior
        .SetSelectionProperty sp_Smoothing, toolpanel_Selections.cboSelSmoothing.ListIndex
        .SetSelectionProperty sp_FeatheringRadius, toolpanel_Selections.sltSelectionFeathering.Value
        If (curShape <> ss_Wand) Then .SetSelectionProperty sp_BorderWidth, toolpanel_Selections.sltSelectionBorder(Selections.GetSelectionSubPanelFromCurrentTool).Value
        .SetSelectionProperty sp_RoundedCornerRadius, toolpanel_Selections.sltCornerRounding.Value
        If (curShape = ss_Polygon) Then .SetSelectionProperty sp_PolygonCurvature, toolpanel_Selections.sltPolygonCurvature.Value
        If (curShape = ss_Lasso) Then .SetSelectionProperty sp_SmoothStroke, toolpanel_Selections.sltSmoothStroke.Value
        If (curShape = ss_Wand) Then
            .SetSelectionProperty sp_WandTolerance, toolpanel_Selections.sltWandTolerance.Value
            .SetSelectionProperty sp_WandSampleMerged, toolpanel_Selections.btsWandMerge.ListIndex
            .SetSelectionProperty sp_WandSearchMode, toolpanel_Selections.btsWandArea.ListIndex
            .SetSelectionProperty sp_WandCompareMethod, toolpanel_Selections.cboWandCompare.ListIndex
        End If
    End With
    
    'Set the first two coordinates of this selection to this mouseclick's location
    PDImages.GetActiveImage.MainSelection.SetInitialCoordinates x, y
    SyncTextToCurrentSelection PDImages.GetActiveImageID()
    PDImages.GetActiveImage.MainSelection.RequestNewMask
    
    'Make the selection tools visible
    SetUIGroupState PDUI_Selections, True
    SetUIGroupState PDUI_SelectionTransforms, True
    
    'Redraw the screen
    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                        
End Sub

'Are selections currently allowed?  Program states like "no open images" prevent selections from being created,
' and individual functions can use this function to determine that state.  Passing TRUE for the
' "transformableMatters" param will add a check for an existing, transformable-type selection (squares, etc)
' to the evaluation list.  (These have their own unique UI requirements.)
Public Function SelectionsAllowed(ByVal transformableMatters As Boolean) As Boolean
    
    SelectionsAllowed = False
    
    If PDImages.IsImageActive() Then
        If PDImages.GetActiveImage.IsSelectionActive And (Not PDImages.GetActiveImage.MainSelection Is Nothing) Then
            If (Not PDImages.GetActiveImage.MainSelection.GetAutoRefreshSuspend()) Then
                If transformableMatters Then
                    SelectionsAllowed = PDImages.GetActiveImage.MainSelection.IsTransformable
                Else
                    SelectionsAllowed = True
                End If
            Else
                Debug.Print "selection refresh suspended"
            End If
        End If
    End If
    
End Function

'Call at program startup.
' At present, all this function does is cache the current user preferences for selection rendering settings.
' This ensures the settings are up-to-date, even if the user does not activate a specific selection tool.
' (Why does this matter? Selections can be loaded directly from file, without ever invoking a tool, so we
' need to ensure rendering settings are up-to-date when the program starts.)
Public Sub InitializeSelectionRendering()

    If UserPrefs.IsReady Then
        
        'Rendering mode (marching ants, highlight, etc)
        m_CurSelectionMode = UserPrefs.GetPref_Long("Tools", "SelectionRenderMode", 0)
        
        'Highlight, lightbox mode render settings
        m_SelHighlightColor = Colors.GetRGBLongFromHex(UserPrefs.GetPref_String("Tools", "SelectionHighlightColor", "#FF3A48"))
        m_SelHighlightOpacity = UserPrefs.GetPref_Float("Tools", "SelectionHighlightOpacity", 50!)
        m_SelLightboxColor = Colors.GetRGBLongFromHex(UserPrefs.GetPref_String("Tools", "SelectionLightboxColor", "#000000"))
        m_SelLightboxOpacity = UserPrefs.GetPref_Float("Tools", "SelectionLightboxOpacity", 50!)
        
    End If

End Sub

'Whenever a selection render setting changes (like switching between outline and highlight mode), you must call this function
' so that we can cache the new render settings.
Public Sub NotifySelectionRenderChange(ByVal settingType As PD_SelectionRenderSetting, ByVal newValue As Variant)
    
    Select Case settingType
        
        Case pdsr_RenderMode
            m_CurSelectionMode = newValue
            
            'Selection rendering settings are cached in PD's main preferences file.  This allows outside functions to access
            ' them correctly, even if selection tools have not been loaded this session.  (This can happen if the user runs
            ' the program, loads an image, then loads a selection directly from file, without invoking a specific tool.)
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionRenderMode", Trim$(Str$(m_CurSelectionMode))
            
        Case pdsr_HighlightColor
            m_SelHighlightColor = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionHighlightColor", Colors.GetHexStringFromRGB(m_SelHighlightColor)
        
        Case pdsr_HighlightOpacity
            m_SelHighlightOpacity = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionHighlightOpacity", m_SelHighlightOpacity
            
        Case pdsr_LightboxColor
            m_SelLightboxColor = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionLightboxColor", Colors.GetHexStringFromRGB(m_SelLightboxColor)
        
        Case pdsr_LightboxOpacity
            m_SelLightboxOpacity = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionlightboxOpacity", m_SelLightboxOpacity
            
    End Select
    
End Sub

Public Function GetSelectionRenderMode() As PD_SelectionRender
    GetSelectionRenderMode = m_CurSelectionMode
End Function

Public Function GetSelectionColor_Highlight() As Long
    GetSelectionColor_Highlight = m_SelHighlightColor
End Function

Public Function GetSelectionOpacity_Highlight() As Single
    GetSelectionOpacity_Highlight = m_SelHighlightOpacity
End Function

Public Function GetSelectionColor_Lightbox() As Long
    GetSelectionColor_Lightbox = m_SelLightboxColor
End Function

Public Function GetSelectionOpacity_Lightbox() As Single
    GetSelectionOpacity_Lightbox = m_SelLightboxOpacity
End Function

'Keypresses on a source canvas are passed here.  The caller doesn't need pass anything except relevant keycodes, and a reference
' to itself (so we can relay canvas modifications).
Public Sub NotifySelectionKeyDown(ByRef srcCanvas As pdCanvas, ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)

    'Handle arrow keys first
    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then

        'If a selection is active, nudge it using the arrow keys
        If (PDImages.GetActiveImage.IsSelectionActive And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster)) Then
            
            Dim canvasUpdateRequired As Boolean
            canvasUpdateRequired = False
            
            'Suspend automatic redraws until all arrow keys have been processed
            srcCanvas.SetRedrawSuspension True
            
            'If scrollbars are visible, nudge the canvas in the direction of the arrows.
            If srcCanvas.GetScrollVisibility(pdo_Vertical) Then
                If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Then canvasUpdateRequired = True
                If (vkCode = VK_UP) Then srcCanvas.SetScrollValue pdo_Vertical, srcCanvas.GetScrollValue(pdo_Vertical) - 1
                If (vkCode = VK_DOWN) Then srcCanvas.SetScrollValue pdo_Vertical, srcCanvas.GetScrollValue(pdo_Vertical) + 1
            End If
            
            If srcCanvas.GetScrollVisibility(pdo_Horizontal) Then
                If (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then canvasUpdateRequired = True
                If (vkCode = VK_LEFT) Then srcCanvas.SetScrollValue pdo_Horizontal, srcCanvas.GetScrollValue(pdo_Horizontal) - 1
                If (vkCode = VK_RIGHT) Then srcCanvas.SetScrollValue pdo_Horizontal, srcCanvas.GetScrollValue(pdo_Horizontal) + 1
            End If
            
            'Re-enable automatic redraws
            srcCanvas.SetRedrawSuspension False
            
            'Redraw the viewport if necessary
            If canvasUpdateRequired Then
                markEventHandled = True
                Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), srcCanvas
            End If
            
        End If
    
    'Handle non-arrow keys here.  (Note: most non-arrow keys are not meant to work with key-repeating,
    ' so they are handled in the KeyUp event instead.)
    Else
        
    End If
    
End Sub

Public Sub NotifySelectionKeyUp(ByRef srcCanvas As pdCanvas, ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)

    'Delete key: if a selection is active, erase the selected area
    If (vkCode = VK_DELETE) And PDImages.GetActiveImage.IsSelectionActive Then
        markEventHandled = True
        Process "Erase selected area", False, BuildParamList("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Layer
    End If
    
    'Escape key: if a selection is active, clear it
    If (vkCode = VK_ESCAPE) And PDImages.GetActiveImage.IsSelectionActive Then
        markEventHandled = True
        Process "Remove selection", , , UNDO_Selection
    End If
    
    'Enter/return keys: for polygon selections, this will close the current selection
    If ((vkCode = VK_RETURN) Or (vkCode = VK_SPACE)) And (g_CurrentTool = SELECT_POLYGON) Then
        
        'A selection must be in-progress
        If PDImages.GetActiveImage.IsSelectionActive Then
        
            'The selection must *not* be closed yet, but there must be enough points to successfully close it
            If (Not PDImages.GetActiveImage.MainSelection.GetPolygonClosedState) And (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 2) Then
            
                'Close the selection
                PDImages.GetActiveImage.MainSelection.SetPolygonClosedState True
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI 0
                
                'Fully process the selection (important when recording macros!)
                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                
                'Redraw the viewport
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            End If
        
        End If
        
    End If
    
    'Backspace key: for lasso and polygon selections, retreat back one or more coordinates, giving the user a chance to
    ' correct any potential mistakes.
    If (vkCode = VK_BACK) And ((g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_POLYGON)) And PDImages.GetActiveImage.IsSelectionActive And (Not PDImages.GetActiveImage.MainSelection.IsLockedIn) Then
        
        markEventHandled = True
        
        'Polygons: do not allow point removal if the polygon has already been successfully closed.
        If (g_CurrentTool = SELECT_POLYGON) Then
            If (Not PDImages.GetActiveImage.MainSelection.GetPolygonClosedState) Then PDImages.GetActiveImage.MainSelection.RemoveLastPolygonPoint
        
        'Lassos: do not allow point removal if the lasso has already been successfully closed.
        Else
        
            If (Not PDImages.GetActiveImage.MainSelection.GetLassoClosedState) Then
        
                'Ask the selection object to retreat its position
                Dim newImageX As Double, newImageY As Double
                PDImages.GetActiveImage.MainSelection.RetreatLassoPosition newImageX, newImageY
                
                'The returned coordinates will be in image coordinates.  Convert them to viewport coordinates.
                Dim newCanvasX As Double, newCanvasY As Double
                Drawing.ConvertImageCoordsToCanvasCoords srcCanvas, PDImages.GetActiveImage(), newImageX, newImageY, newCanvasX, newCanvasY
                
                'Finally, convert the canvas coordinates to screen coordinates, and move the cursor accordingly
                srcCanvas.SetCursorToCanvasPosition newCanvasX, newCanvasY
                
            End If
            
        End If
        
        'Redraw the screen to reflect this new change.
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
    
    End If
                
End Sub

Public Sub NotifySelectionMouseDown(ByRef srcCanvas As pdCanvas, ByVal imgX As Single, ByVal imgY As Single)
    
    'Before processing the mouse event, check to see if a selection is already active.  If it is, and its type
    ' does *not* match the current selection tool, invalidate the old selection and apply the new type before proceeding.
    If PDImages.GetActiveImage.IsSelectionActive Then
        If (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> Selections.GetSelectionShapeFromCurrentTool()) Then
            PDImages.GetActiveImage.SetSelectionActive False
            PDImages.GetActiveImage.MainSelection.SetSelectionShape Selections.GetSelectionShapeFromCurrentTool()
        End If
    End If
        
    'Because the wand tool is extremely simple, handle it specially
    If (g_CurrentTool = SELECT_WAND) Then
    
        'Magic wand selections never transform - they only generate anew
        Selections.InitSelectionByPoint imgX, imgY
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
        
    Else
        
        'Check to see if a selection is already active.  If it is, see if the user is allowed to transform it.
        If PDImages.GetActiveImage.IsSelectionActive Then
        
            'Check the mouse coordinates of this click.
            Dim sCheck As PD_PointOfInterest
            sCheck = Selections.IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
            
            'If a point of interest was clicked, initiate a transform
            If (sCheck <> poi_Undefined) And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Polygon) And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster) Then
                
                'Initialize a selection transformation
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI sCheck
                PDImages.GetActiveImage.MainSelection.SetInitialTransformCoordinates imgX, imgY
                                
            'If a point of interest was *not* clicked, erase any existing selection and start a new one
            Else
                
                'Polygon selections require special handling, because they don't operate on the "mouse up = complete" assumption.
                ' They are completed when the user re-clicks the first point.  Any clicks prior to that point are treated as
                ' an instruction to add a new points.
                If (g_CurrentTool = SELECT_POLYGON) Then
                    
                    'First, see if the selection is locked in.  If it is, treat this is a regular transformation.
                    If PDImages.GetActiveImage.MainSelection.IsLockedIn Then
                        PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI sCheck
                        PDImages.GetActiveImage.MainSelection.SetInitialTransformCoordinates imgX, imgY
                    
                    'Selection is not locked in, meaning the user is still constructing it.
                    Else
                    
                        'If the user clicked on the initial polygon point, attempt to close the polygon
                        If (sCheck = 0) And (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 2) Then
                            PDImages.GetActiveImage.MainSelection.SetPolygonClosedState True
                            PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI 0
                        
                        'The user did not click the initial polygon point, meaning we should add this coordinate as a new polygon point.
                        Else
                            
                            'Remove the current transformation mode (if any)
                            PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI poi_Undefined
                            PDImages.GetActiveImage.MainSelection.OverrideTransformMode False
                            
                            'Add the new point
                            If (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints = 0) Then
                                Selections.InitSelectionByPoint imgX, imgY
                            Else
                                
                                If (sCheck = poi_Undefined) Or (sCheck = poi_Interior) Then
                                    PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                                    PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints - 1
                                Else
                                    PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI sCheck
                                End If
                                
                            End If
                            
                            'Reinstate transformation mode, using the index of the new point as the transform ID
                            PDImages.GetActiveImage.MainSelection.SetInitialTransformCoordinates imgX, imgY
                            PDImages.GetActiveImage.MainSelection.OverrideTransformMode True
                            
                            'Redraw the screen
                            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                            
                        End If
                    
                    End If
                    
                Else
                    Selections.InitSelectionByPoint imgX, imgY
                End If
                
            End If
        
        'If a selection is not active, start a new one
        Else
            
            Selections.InitSelectionByPoint imgX, imgY
            
            'Polygon selections require special handling, as usual.  After creating the initial point, we want to immediately initiate
            ' transform mode, because dragging the mouse will simply move the newly created point.
            If (g_CurrentTool = SELECT_POLYGON) Then
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints - 1
                PDImages.GetActiveImage.MainSelection.OverrideTransformMode True
            End If
            
        End If
        
    End If
    
End Sub

'The only selection tool that responds to double-click events is the polygon selection tool.
' Photoshop convention (mirrored by GIMP, Krita) is to close the polygon on a double-click.
Public Sub NotifySelectionMouseDblClick(ByRef srcCanvas As pdCanvas, ByVal imgX As Single, ByVal imgY As Single)
    
    'Polygon selections only
    If (g_CurrentTool = SELECT_POLYGON) Then
    
        'A selection must be in-progress
        If PDImages.GetActiveImage.IsSelectionActive Then
        
            'The selection must *not* be closed yet
            If (Not PDImages.GetActiveImage.MainSelection.GetPolygonClosedState) And (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 2) Then
                
                'Set a flag to note that a double-click just occurred.  (See notes at the
                ' top of this module for details.)
                m_DblClickOccurred = True
                
                'Remove the last point (the point created by the first click of this
                ' double-click event), but *only* if there are enough valid points
                ' to create a polygon selection without it!
                If (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 3) Then PDImages.GetActiveImage.MainSelection.RemoveLastPolygonPoint
                
                'Close the selection and make the first point the active one
                PDImages.GetActiveImage.MainSelection.SetPolygonClosedState True
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI 0
                
                'Fully process the selection (important when recording macros!)
                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                
                'Redraw the viewport
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            End If
        
        End If
    
    End If

End Sub

Public Sub NotifySelectionMouseLeave(ByRef srcCanvas As pdCanvas)

    'When the polygon selection tool is being used, redraw the canvas when the mouse leaves
    If (g_CurrentTool = SELECT_POLYGON) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas

End Sub

Public Sub NotifySelectionMouseMove(ByRef srcCanvas As pdCanvas, ByVal lmbState As Boolean, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single, ByVal numOfCanvasMoveEvents As Long)
    
    'Handling varies based on the current mouse state, obviously.
    If lmbState Then
        
        'Basic selection tools
        Select Case g_CurrentTool
            
            Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON
                
                'First, check to see if a selection is both active and transformable.
                If PDImages.GetActiveImage.IsSelectionActive And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster) Then
                    
                    'If the SHIFT key is down, notify the selection engine that a square shape is requested
                    PDImages.GetActiveImage.MainSelection.RequestSquare (Shift And vbShiftMask)
                    
                    'Pass new points to the active selection
                    PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                    Selections.SyncTextToCurrentSelection PDImages.GetActiveImageID()
                                        
                End If
                
                'Force a redraw of the viewport
                If (numOfCanvasMoveEvents > 1) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            
            'Lasso selections are handled specially, because mouse move events control the drawing of the lasso
            Case SELECT_LASSO
            
                'First, check to see if a selection is active
                If PDImages.GetActiveImage.IsSelectionActive Then
                    
                    'Pass new points to the active selection
                    PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                                        
                End If
                
                'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                ' while in debug mode.
                If UserPrefs.GenerateDebugLogs Then
                    Message "Release the mouse button to complete the lasso selection", "DONOTLOG"
                Else
                    Message "Release the mouse button to complete the lasso selection"
                End If
                
                'Force a redraw of the viewport
                If (numOfCanvasMoveEvents > 1) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            
            'Wand selections are easier than other selection types, because they don't support any special transforms
            Case SELECT_WAND
                If PDImages.GetActiveImage.IsSelectionActive Then
                    PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                End If
        
        End Select
    
    'The left mouse button is *not* down
    Else
        
        'Notify the selection of the currently hovered point of interest, if any
        Dim selPOI As PD_PointOfInterest
        selPOI = Selections.IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
        
        If (selPOI <> PDImages.GetActiveImage.MainSelection.GetActiveSelectionPOI(False)) Then
            PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI selPOI
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), srcCanvas
        Else
            If (g_CurrentTool = SELECT_POLYGON) Then If (g_CurrentTool = SELECT_POLYGON) Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), srcCanvas
        End If
        
    End If
        
End Sub

Public Sub NotifySelectionMouseUp(ByRef srcCanvas As pdCanvas, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single, ByVal clickEventAlsoFiring As Boolean, ByVal wasSelectionActiveBeforeMouseEvents As Boolean)
    
    'If a double-click just occurred, reset the flag and exit - do NOT process this click further
    If m_DblClickOccurred Then
        m_DblClickOccurred = False
        Exit Sub
    End If
    
    Dim eraseThisSelection As Boolean
    
    Select Case g_CurrentTool
    
        'Most selection tools are handled identically
        Case SELECT_RECT, SELECT_CIRC, SELECT_LASSO
        
            'If a selection was being drawn, lock it into place
            If PDImages.GetActiveImage.IsSelectionActive Then
                
                'Check to see if this mouse location is the same as the initial mouse press. If it is, and that particular
                ' point falls outside the selection, clear the selection from the image.
                Dim selBounds As RectF
                selBounds = PDImages.GetActiveImage.MainSelection.GetCornersLockedRect
                
                eraseThisSelection = (clickEventAlsoFiring And (IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage()) = -1))
                If (Not eraseThisSelection) Then eraseThisSelection = ((selBounds.Width <= 0) And (selBounds.Height <= 0))
                
                If eraseThisSelection Then
                    Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                    
                'The mouse is being released after a significant move event, or on a point of interest to the current selection.
                Else
                
                    'If the selection is not raster-type, pass these final mouse coordinates to it
                    If (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster) Then
                        PDImages.GetActiveImage.MainSelection.RequestSquare (Shift And vbShiftMask)
                        PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                        SyncTextToCurrentSelection PDImages.GetActiveImageID()
                    End If
                
                    'Check to see if all selection coordinates are invalid (e.g. off-image).  If they are, forget about this selection.
                    If PDImages.GetActiveImage.MainSelection.AreAllCoordinatesInvalid Then
                        Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                    Else
                        
                        'Depending on the type of transformation that may or may not have been applied, call the appropriate processor function.
                        ' This is required to add the current selection event to the Undo/Redo chain.
                        If (g_CurrentTool = SELECT_LASSO) Then
                        
                            'Creating a new selection
                            If (PDImages.GetActiveImage.MainSelection.GetActiveSelectionPOI = poi_Undefined) Then
                                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            
                            'Moving an existing selection
                            Else
                                Process "Move selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            End If
                        
                        'All other selection types use identical transform identifiers
                        Else
                        
                            Dim transformType As PD_PointOfInterest
                            transformType = PDImages.GetActiveImage.MainSelection.GetActiveSelectionPOI
                            
                            'Creating a new selection
                            If (transformType = poi_Undefined) Then
                                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            
                            'Moving an existing selection
                            ElseIf (transformType = 8) Then
                                Process "Move selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                                
                            'Anything else is assumed to be resizing an existing selection
                            Else
                                Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                                        
                            End If
                        
                        End If
                        
                    End If
                    
                End If
                
                'Creating a brand new selection always necessitates a redraw of the current canvas
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            'If the selection is not active, make sure it stays that way
            Else
                PDImages.GetActiveImage.MainSelection.LockRelease
            End If
            
            'Synchronize the selection text box values with the final selection
            Selections.SyncTextToCurrentSelection PDImages.GetActiveImageID()
            
        
        'As usual, polygon selections require special considerations.
        Case SELECT_POLYGON
        
            'If a selection was being drawn, lock it into place
            If PDImages.GetActiveImage.IsSelectionActive Then
            
                'Check to see if the selection is already locked in.  If it is, we need to check for an "erase selection" click.
                eraseThisSelection = PDImages.GetActiveImage.MainSelection.GetPolygonClosedState And clickEventAlsoFiring
                eraseThisSelection = eraseThisSelection And (IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage()) = -1)
                
                If eraseThisSelection Then
                    Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                
                Else
                    
                    'If the polygon is already closed, we want to lock in the newly modified polygon
                    If PDImages.GetActiveImage.MainSelection.GetPolygonClosedState Then
                        
                        'Polygons use a different transform numbering convention than other selection tools, because the number
                        ' of points involved aren't fixed.
                        Dim polyPoint As Long
                        polyPoint = Selections.IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
                        
                        'Move selection
                        If (polyPoint = PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints) Then
                            Process "Move selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                        
                        'Create OR resize, depending on whether the initial point is being clicked for the first time, or whether
                        ' it's being click-moved
                        ElseIf (polyPoint = 0) Then
                            If clickEventAlsoFiring Then
                                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            Else
                                Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            End If
                                
                        'No point of interest means this click lies off-image; this could be a "clear selection" event (if a Click
                        ' event is also firing), or a "move polygon point" event (if the user dragged a point off-image).
                        ElseIf (polyPoint = -1) Then
                            
                            'If the user has clicked a blank spot unrelated to the selection, we want to remove the active selection
                            If clickEventAlsoFiring Then
                                Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                                
                            'If they haven't clicked, this could simply indicate that they dragged a polygon point off the polygon
                            ' and into some new region of the image.
                            Else
                                PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                                Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            End If
                            
                        'Anything else is a resize
                        Else
                            Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                        End If
                        
                        'After all that work, we want to perform one final check to see if all selection coordinates are invalid
                        ' (e.g. if they all lie off-image, which can happen if the user drags all polygon points off-image).
                        ' If they are, we're going to erase this selection, as it's invalid.
                        eraseThisSelection = PDImages.GetActiveImage.MainSelection.IsLockedIn And PDImages.GetActiveImage.MainSelection.AreAllCoordinatesInvalid
                        If eraseThisSelection Then Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                        
                    'If the polygon is *not* closed, we want to add this as a new polygon point
                    Else
                    
                        'Pass these final mouse coordinates to the selection engine
                        PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                        
                        'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                        ' while in debug mode.
                        If (Not wasSelectionActiveBeforeMouseEvents) Then
                            If UserPrefs.GenerateDebugLogs Then
                                Message "Click on the first point to complete the polygon selection", "DONOTLOG"
                            Else
                                Message "Click on the first point to complete the polygon selection"
                            End If
                        End If
                        
                    End If
                
                'End erase vs create check
                End If
                
                'After all selection settings have been applied, forcibly redraw the source canvas
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            
            '(Failsafe check) - if a selection is not active, make sure it stays that way
            Else
                PDImages.GetActiveImage.MainSelection.LockRelease
            End If
            
        'Magic wand selections are actually the easiest to handle, as they don't really support post-creation transforms
        Case SELECT_WAND
            
            'Failsafe check for active selections
            If PDImages.GetActiveImage.IsSelectionActive Then
                
                'Supply the final coordinates to the selection engine (as the user may be dragging around the active point)
                PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                
                'Check to see if all selection coordinates are invalid (e.g. off-image).
                ' - If they are, forget about this selection.
                ' - If they are not, commit this selection permanently
                eraseThisSelection = PDImages.GetActiveImage.MainSelection.AreAllCoordinatesInvalid
                If eraseThisSelection Then
                    Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                Else
                    Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                End If
                
                'Force a redraw of the screen
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            'Failsafe check for inactive selections
            Else
                PDImages.GetActiveImage.MainSelection.LockRelease
            End If
            
    End Select
    
End Sub
