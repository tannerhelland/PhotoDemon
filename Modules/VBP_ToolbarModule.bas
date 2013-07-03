Attribute VB_Name = "Toolbar"
'***************************************************************************
'Toolbar Interface
'Copyright ©2001-2013 by Tanner Helland
'Created: 4/15/01
'Last updated: 03/July/13
'Last update: added tSelectionTransform for more fine-grained selection activation/deactivation.  Transformable selections must also enable
'             the coordinate text boxes on the main form, while non-transformable ones must deactivate these.
'
'Module for enabling/disabling toolbar buttons and menus.  Note that the toolbar was removed in June '12 in favor of
' the new left-hand bar; this module remains, however, because the code handles menu items and the left-hand bar
' (not just toolbar-related code, regardless of what the name implies :).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

Public Const tOpen As Long = 0
Public Const tSave As Long = 1
Public Const tSaveAs As Long = 2
Public Const tCopy As Long = 3
Public Const tPaste As Long = 4
Public Const tUndo As Long = 5
Public Const tImageOps As Long = 6
Public Const tFilter As Long = 7
Public Const tRedo As Long = 8
'Public Const tHistogram As Long = 9
Public Const tMacro As Long = 10
Public Const tEdit As Long = 11
Public Const tRepeatLast As Long = 12
Public Const tSelection As Long = 13
Public Const tSelectionTransform As Long = 14
Public Const tImgMode32bpp As Long = 15
Public Const tMetadata As Long = 16
Public Const tGPSMetadata As Long = 17


'tInit enables or disables a specified button and/or menu item
Public Sub tInit(tButton As Long, tState As Boolean)
    
    Dim i As Long
    
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
            If tState Then
                FormMain.cmdUndo.ToolTip = pdImages(CurrentImage).getUndoProcessID ' GetNameOfProcess(pdImages(CurrentImage).getUndoProcessID)
                FormMain.MnuUndo.Caption = g_Language.TranslateMessage("Undo:") & " " & pdImages(CurrentImage).getUndoProcessID 'GetNameOfProcess(pdImages(CurrentImage).getUndoProcessID)
                ResetMenuIcons
            Else
                FormMain.cmdUndo.ToolTip = ""
                FormMain.MnuUndo.Caption = g_Language.TranslateMessage("Undo")
                ResetMenuIcons
            End If
            
        'ImageOps is all Image-related menu items; it enables/disables the Image, Select, Color, View (most items), and Print menus
        Case tImageOps
            If FormMain.MnuImageTop.Enabled <> tState Then
                FormMain.MnuImageTop.Enabled = tState
                'Use this same command to disable other menus
                FormMain.MnuPrint.Enabled = tState
                FormMain.MnuFitOnScreen.Enabled = tState
                FormMain.MnuFitWindowToImage.Enabled = tState
                FormMain.MnuZoomIn.Enabled = tState
                FormMain.MnuZoomOut.Enabled = tState
                FormMain.MnuSelectTop.Enabled = tState
                FormMain.MnuColorTop.Enabled = tState
                FormMain.MnuWindowTop.Enabled = tState
                
                For i = 0 To FormMain.MnuSpecificZoom.Count - 1
                    FormMain.MnuSpecificZoom(i).Enabled = tState
                Next i
                
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
                FormMain.cmdRedo.ToolTip = pdImages(CurrentImage).getRedoProcessID 'GetNameOfProcess(pdImages(CurrentImage).getRedoProcessID)
                FormMain.MnuRedo.Caption = g_Language.TranslateMessage("Redo:") & " " & pdImages(CurrentImage).getRedoProcessID 'GetNameOfProcess(pdImages(CurrentImage).getRedoProcessID) '& vbTab & "Ctrl+Alt+Z"
                ResetMenuIcons
            Else
                FormMain.cmdRedo.ToolTip = ""
                FormMain.MnuRedo.Caption = g_Language.TranslateMessage("Redo")
                ResetMenuIcons
            End If
            
        'Macro (top-level menu)
        Case tMacro
            If FormMain.mnuTool(2).Enabled <> tState Then FormMain.mnuTool(2).Enabled = tState
        
        'Edit (top-level menu)
        Case tEdit
            If FormMain.MnuEdit.Enabled <> tState Then FormMain.MnuEdit.Enabled = tState
        
        'Repeat last action (menu item only)
        Case tRepeatLast
            If FormMain.MnuRepeatLast.Enabled <> tState Then FormMain.MnuRepeatLast.Enabled = tState
            
        'Selections in general
        Case tSelection
            
            'If selections are not active, clear all the selection value textboxes
            If Not tState Then
                For i = 0 To FormMain.tudSel.Count - 1
                    FormMain.tudSel(i).Value = 0
                Next i
            End If
            
            'Set selection text boxes (only the location ones!) to enable only when a selection is active.  Other selection controls can
            ' remain active even without a selection present; this allows the user to set certain parameters in advance, so when they
            ' actually draw a selection, it already has the attributes they want.
            For i = 0 To FormMain.tudSel.Count - 1
                If FormMain.tudSel(i).Enabled <> tState Then FormMain.tudSel(i).Enabled = tState
            Next i
                                    
            'Selection enabling/disabling also affects the Crop to Selection command
            If FormMain.MnuImage(7).Enabled <> tState Then FormMain.MnuImage(7).Enabled = tState
        
        'Transformable selection controls specifically
        Case tSelectionTransform
        
            'Set selection text boxes (only the location ones!) to enable only when a selection is active.  Other selection controls can
            ' remain active even without a selection present; this allows the user to set certain parameters in advance, so when they
            ' actually draw a selection, it already has the attributes they want.
            For i = 0 To FormMain.tudSel.Count - 1
                If FormMain.tudSel(i).Enabled <> tState Then FormMain.tudSel(i).Enabled = tState
            Next i
        
        '32bpp color mode (e.g. add/remove alpha channel)
        Case tImgMode32bpp
            
            FormMain.MnuTransparency(0).Enabled = Not tState
            FormMain.MnuTransparency(1).Enabled = tState
            
        Case tMetadata
        
            If FormMain.MnuMetadata(0).Enabled <> tState Then FormMain.MnuMetadata(0).Enabled = tState
        
        Case tGPSMetadata
        
            If FormMain.MnuMetadata(2).Enabled <> tState Then FormMain.MnuMetadata(2).Enabled = tState
            
    End Select
    
End Sub
