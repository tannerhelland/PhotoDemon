VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMRURecentMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'RecentMacros MRU List
'Copyright 2005-2015 by Tanner Helland
'Created: 17/February/15
'Last updated: 17/February/15
'Last updated by: Raj
'Last update:
'
'This class is a specific implementation of the IMRUList Interface.
'  It works with pdMRUManager to create and maintain the Recent
'  Macros list.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit
Implements IMRUList

Private Sub IMRUList_BeforeListSave(manager As pdMRUManager, Optional cancel As Boolean = False)

End Sub

Private Sub IMRUList_BeforeNewFileAdded(manager As pdMRUManager, newFile As String, Optional cancel As Boolean = False)

End Sub

Private Property Get IMRUList_SupportsThumbNails() As Boolean
    IMRUList_SupportsThumbNails = False
End Property

Private Function IMRUList_GetThumbNailPath(manager As pdMRUManager, Index As Long) As String
    IMRUList_GetThumbNailPath = ""
End Function

Private Sub IMRUList_SaveThumbNail(filePath As String, srcImage As pdImage)
    
End Sub


Private Sub IMRUList_AfterListCleared()
    'Unload all recent file menu entries
    If FormMain.mnuRecentMacros.Count > 1 Then
        Dim i As Long
        For i = FormMain.mnuRecentMacros.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecentMacros(i)
        Next i
    End If
End Sub

Private Sub IMRUList_AfterListSaved(manager As pdMRUManager)

End Sub

Private Sub IMRUList_AfterNewFileAdded(manager As pdMRUManager, newFile As String, mruCaption As String)
    
    'Redraw the MRU menu based on the updated list
    If Not FormMain.mnuRecentMacros(0).Enabled Then
        FormMain.mnuRecentMacros(0).Enabled = True
        FormMain.mnuClearRecentMacros.Enabled = True
    End If
    
    'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
    FormMain.mnuRecentMacros(0).Caption = mruCaption
        
    Dim i As Long
    Dim numEntries As Long
    numEntries = manager.MRU_ReturnCount()
    'If there are more entries on the list besides this one, update their menus now.
    If numEntries > 1 Then
    
        'Unload existing menus...
        For i = FormMain.mnuRecentMacros.Count - 1 To 1 Step -1
            Unload FormMain.mnuRecentMacros(i)
        Next i
        
        'Load new menus...
        For i = 1 To numEntries - 1
        
            Load FormMain.mnuRecentMacros(i)
            
            'Based on the user's preference, display just the filename or the entire file path (up to the max character length)
            FormMain.mnuRecentMacros(i).Caption = manager.getSpecificMRUCaption(i)
                        
        Next i
    End If
End Sub

Private Sub IMRUList_BeforeListLoad(manager As pdMRUManager, Optional ByRef cancel As Boolean = False)

End Sub



Private Sub IMRUList_OnItemLoad(manager As pdMRUManager, Index As Long, mruCaption As String)
    'Note that we do not need to load mnuRecDocs(0) - it is always loaded.  Only load menu items greater than this.
    If Index > 0 Then
        Load FormMain.mnuRecentMacros(Index)
    Else
        FormMain.mnuRecentMacros(Index).Enabled = True
    End If
    
    FormMain.mnuRecentMacros(Index).Caption = mruCaption
    
End Sub

Private Sub IMRUList_AfterListLoaded(manager As pdMRUManager)
    'Make sure the "Clear MRU" option is visible and enabled
    If manager.MRU_ReturnCount > 0 Then
        FormMain.mnuClearRecentMacros.Enabled = True
    Else
        FormMain.mnuClearRecentMacros.Enabled = False
    End If
End Sub

Private Sub IMRUList_OnListIsEmpty()
    FormMain.mnuRecentMacros(0).Caption = g_Language.TranslateMessage("Empty")
    FormMain.mnuRecentMacros(0).Enabled = False
    FormMain.mnuClearRecentMacros.Enabled = False
    'FormMain.MnuRecentSepBar1.Visible = False
    'FormMain.MnuClearMRU.Visible = False
    'FormMain.MnuLoadAllMRU.Visible = False
End Sub

Private Sub IMRUList_OnMaxLimitChanged(manager As pdMRUManager)
    
End Sub


Private Sub IMRUList_TrimListAtPosition(manager As pdMRUManager, Index As Long)

End Sub

Private Property Get IMRUList_XMLFileName() As String
    IMRUList_XMLFileName = g_UserPreferences.getPresetPath & "Program_RecentMacros.xml"
End Property



