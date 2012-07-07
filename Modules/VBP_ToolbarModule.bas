Attribute VB_Name = "Toolbar"
'***************************************************************************
'Toolbar Interface
'©2000-2012 Tanner Helland
'Created: 4/15/01
'Last updated: 15/June/12
'Last update: because PhotoDemon can now load multiple images simultaneously, the toolbar would flicker madly
'             as various buttons were enabled/disabled upon each new form's creation.  Now, those button states
'             are checked against requests to enable/disable, and a change is made only when absolutely necessary.
'
'Module for enabling/disabling toolbar buttons and menus.
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
Public Const tHistogram As Byte = 9
Public Const tMacro As Byte = 10
Public Const tEdit As Byte = 11
Public Const tRepeatLast As Byte = 12

Public Sub tInit(tButton As Byte, tState As Boolean)
    
    Select Case tButton
        Case tOpen
            If FormMain.MnuOpen.Enabled <> tState Then
                FormMain.cmdOpen.Enabled = tState
                FormMain.MnuOpen.Enabled = tState
            End If
        Case tSave
            If FormMain.MnuSave.Enabled <> tState Then
                FormMain.cmdSave.Enabled = tState
                FormMain.MnuSave.Enabled = tState
            End If
        Case tSaveAs
            If FormMain.MnuSaveAs.Enabled <> tState Then FormMain.MnuSaveAs.Enabled = tState
        Case tCopy
            If FormMain.MnuCopy.Enabled <> tState Then FormMain.MnuCopy.Enabled = tState
        Case tPaste
            If FormMain.MnuPaste.Enabled <> tState Then FormMain.MnuPaste.Enabled = tState
        Case tUndo
            If FormMain.MnuUndo.Enabled <> tState Then
                FormMain.cmdUndo.Enabled = tState
                FormMain.MnuUndo.Enabled = tState
            End If
        Case tImageOps
            If FormMain.MnuImage.Enabled <> tState Then
                FormMain.MnuImage.Enabled = tState
                'Cheat and use the same command to disable the color menu...hey, at least it works
                FormMain.MnuColor.Enabled = tState
                'Cheat again and enable/disable the Print menu too
                FormMain.MnuPrint.Enabled = tState
            End If
        Case tFilter
            If FormMain.MnuFilter.Enabled <> tState Then FormMain.MnuFilter.Enabled = tState
        Case tRedo
            If FormMain.MnuRedo.Enabled <> tState Then
                FormMain.cmdRedo.Enabled = tState
                FormMain.MnuRedo.Enabled = tState
            End If
        Case tHistogram
            'If FormMain.MnuHistogramTop.Enabled <> tState Then
            '    FormMain.Toolbar1.Buttons(10).Enabled = tState
            '    changeMade = True
            'End If
            'FormMain.Toolbar1.Buttons(10).Enabled = tState
        Case tMacro
            If FormMain.MnuMacro.Enabled <> tState Then FormMain.MnuMacro.Enabled = tState
        Case tEdit
            If FormMain.MnuEdit.Enabled <> tState Then FormMain.MnuEdit.Enabled = tState
        Case tRepeatLast
            If FormMain.MnuRepeatLast.Enabled <> tState Then FormMain.MnuRepeatLast.Enabled = tState
    End Select
    
End Sub
