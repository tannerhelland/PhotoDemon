Attribute VB_Name = "Toolbar"
'***************************************************************************
'Toolbar Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 03/October/12
'Last update: Added capability for activating/deactivating the Selection tool interface
'
'Module for enabling/disabling toolbar buttons and menus.  Note that the toolbar was removed in June 2012 in favor of
' the new left-hand bar; this module remains, however, because the code handles menu items and the left-hand bar
' (not just toolbar-related code, regardless of what the name implies :).
'
'***************************************************************************

Option Explicit

Public Const tOpen As Byte = 0
Public Const tSave As Byte = 1
Public Const tSaveAs As Byte = 2
Public Const tCopy As Byte = 3
Public Const tPaste As Byte = 4
Public Const tUndo As Byte = 5
Public Const tImageOps As Byte = 6
Public Const tFilter As Byte = 7
Public Const tRedo As Byte = 8
'Public Const tHistogram As Byte = 9
Public Const tMacro As Byte = 10
Public Const tEdit As Byte = 11
Public Const tRepeatLast As Byte = 12
Public Const tSelection As Byte = 13

'tInit enables or disables a specified button and/or menu item
Public Sub tInit(tButton As Byte, tState As Boolean)
    
    Select Case tButton
        
        'Open (left-hand panel button AND menu item)
        Case tOpen
            If FormMain.MnuOpen.Enabled <> tState Then
                FormMain.cmdOpen.Enabled = tState
                FormMain.MnuOpen.Enabled = tState
            End If
            
        'Save (left-hand panel button AND menu item)
        Case tSave
            If FormMain.MnuSave.Enabled <> tState Then
                FormMain.cmdSave.Enabled = tState
                FormMain.MnuSave.Enabled = tState
            End If
            
        'Save As (menu item only)
        Case tSaveAs
            If FormMain.MnuSaveAs.Enabled <> tState Then
                FormMain.cmdSaveAs.Enabled = tState
                FormMain.MnuSaveAs.Enabled = tState
            End If
        
        'Copy (menu item only)
        Case tCopy
            If FormMain.MnuCopy.Enabled <> tState Then FormMain.MnuCopy.Enabled = tState
        
        'Paste (menu item only)
        Case tPaste
            If FormMain.MnuPaste.Enabled <> tState Then FormMain.MnuPaste.Enabled = tState
        
        'Undo (left-hand panel button AND menu item)
        Case tUndo
            If FormMain.MnuUndo.Enabled <> tState Then
                FormMain.cmdUndo.Enabled = tState
                FormMain.MnuUndo.Enabled = tState
            End If
            'If Undo is being enabled, change the text to match the relevant action that created this Undo file
            If tState = True Then
                FormMain.cmdUndo.ToolTip = GetNameOfProcess(pdImages(CurrentImage).getUndoProcessID)
                FormMain.MnuUndo.Caption = "Undo: " & GetNameOfProcess(pdImages(CurrentImage).getUndoProcessID)
                ResetMenuIcons
            Else
                FormMain.cmdUndo.ToolTip = ""
                FormMain.MnuUndo.Caption = "Undo"
                ResetMenuIcons
            End If
            
        'ImageOps is all Image-related menu items; it enables/disables the Image, Color, View, and Print menus
        Case tImageOps
            If FormMain.MnuImage.Enabled <> tState Then
                FormMain.MnuImage.Enabled = tState
                'Cheat and use the same command to disable the color menu...hey, at least it works
                FormMain.MnuColor.Enabled = tState
                'Cheat again and enable/disable the Print menu
                FormMain.MnuPrint.Enabled = tState
                'Cheat again and enable/disable the View menu
                FormMain.MnuView.Enabled = tState
            End If
        
        'Filter (top-level menu)
        Case tFilter
            If FormMain.MnuFilter.Enabled <> tState Then FormMain.MnuFilter.Enabled = tState
        
        'Redo (left-hand panel button AND menu item)
        Case tRedo
            If FormMain.MnuRedo.Enabled <> tState Then
                FormMain.cmdRedo.Enabled = tState
                FormMain.MnuRedo.Enabled = tState
            End If
            
            'If Redo is being enabled, change the menu text to match the relevant action that created this Undo file
            If tState = True Then
                FormMain.cmdRedo.ToolTip = GetNameOfProcess(pdImages(CurrentImage).getRedoProcessID)
                FormMain.MnuRedo.Caption = "Redo: " & GetNameOfProcess(pdImages(CurrentImage).getRedoProcessID) '& vbTab & "Ctrl+Alt+Z"
                ResetMenuIcons
            Else
                FormMain.cmdRedo.ToolTip = ""
                FormMain.MnuRedo.Caption = "Redo" '& vbTab & "Ctrl+Alt+Z"
                ResetMenuIcons
            End If
            
        'Macro (top-level menu)
        Case tMacro
            If FormMain.MnuMacro.Enabled <> tState Then FormMain.MnuMacro.Enabled = tState
        
        'Edit (top-level menu)
        Case tEdit
            If FormMain.MnuEdit.Enabled <> tState Then FormMain.MnuEdit.Enabled = tState
        
        'Repeat last action (menu item only)
        Case tRepeatLast
            If FormMain.MnuRepeatLast.Enabled <> tState Then FormMain.MnuRepeatLast.Enabled = tState
            
        'Selections
        Case tSelection
            If FormMain.txtSelLeft.Visible <> tState Then FormMain.txtSelLeft.Visible = tState
            If FormMain.txtSelTop.Visible <> tState Then FormMain.txtSelTop.Visible = tState
            If FormMain.txtSelWidth.Visible <> tState Then FormMain.txtSelWidth.Visible = tState
            If FormMain.txtSelHeight.Visible <> tState Then FormMain.txtSelHeight.Visible = tState
            If FormMain.vsSelLeft.Visible <> tState Then FormMain.vsSelLeft.Visible = tState
            If FormMain.vsSelTop.Visible <> tState Then FormMain.vsSelTop.Visible = tState
            If FormMain.vsSelWidth.Visible <> tState Then FormMain.vsSelWidth.Visible = tState
            If FormMain.vsSelHeight.Visible <> tState Then FormMain.vsSelHeight.Visible = tState
            If FormMain.lblSelSize.Visible <> tState Then FormMain.lblSelSize.Visible = tState
            If FormMain.lblSelPosition.Visible <> tState Then FormMain.lblSelPosition.Visible = tState
            
            'Selection enabling/disabling also affects the Crop to Selection command
            If FormMain.MnuCropSelection.Enabled <> tState Then FormMain.MnuCropSelection.Enabled = tState
            
    End Select
    
End Sub
