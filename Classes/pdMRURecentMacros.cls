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
'Recent Files (MRU) handler for PD macros (Tools > Macros)
'Copyright 2005-2022 by Raj Chaudhuri and Tanner Helland
'Created: 17/February/15
'Last updated: 15/August/17
'Last update: implemented support for Unicode menu captions
'Dependencies: - everything listed in IMRUList
'              - Files module (for Unicode support), Crypto module (for fast hashing)
'
'This class is a specific implementation of the IMRUList Interface.  It works alongside pdMRUManager to manage
' the UI for PD's macro recent file list (Tools > Macros menu).
'
'In 2015, Raj Chaudhuri kindly rewrote this class as a replacement for PD's previous pdRecentFiles class.
' Data written by the old class is still compatible with this implementation.  Many thanks to Raj for this
' great improvement!  (Thanks also to Raj for rewriting this class as an implementation of IMRUList, so PD can
' now support recent file lists in multiple places throughout the program.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit
Implements IMRUList


'Macros do not currently support menu thumbnails
Private Property Get IMRUList_SupportsThumbnails() As Boolean
    IMRUList_SupportsThumbnails = False
End Property

Private Function IMRUList_GetThumbnailPath(ByRef mruManager As pdMRUManager, ByVal Index As Long) As String
    IMRUList_GetThumbnailPath = vbNullString
End Function

'The location of the XML file is hard-coded against the current preset path
Private Property Get IMRUList_XMLFileName() As String
    IMRUList_XMLFileName = UserPrefs.GetPresetPath & "Program_RecentMacros.xml"
End Property

'When the recent macro list is cleared, we can unload the entire mnuRecentMacros collection (except index 0, which is simply hidden)
Private Sub IMRUList_AfterListCleared()

    'Unload all recent file menu entries
    If (FormMain.MnuRecentMacros.Count > 1) Then
        
        Dim i As Long
        For i = FormMain.MnuRecentMacros.Count - 1 To 1 Step -1
            Unload FormMain.MnuRecentMacros(i)
        Next i
        
    End If
    
    'Because the "clear list" menu entry sits at the bottom, we must refresh icons when the list is modified
    IconsAndCursors.ResetMenuIcons
    
End Sub

'When a new file is added, we unload the menu collection, then reload it in the proper order
Private Sub IMRUList_AfterFileAdded(ByRef mruManager As pdMRUManager, ByVal newFilePath As String, ByVal mruCaption As String)
    
    'Redraw the MRU menu based on the updated list
    If (Not FormMain.MnuRecentMacros(0).Enabled) Then
        FormMain.MnuRecentMacros(0).Enabled = True
        FormMain.MnuClearRecentMacros.Enabled = True
    End If
    
    'Ask the manager object how many entries it currently maintains
    Dim numEntries As Long
    numEntries = mruManager.MRU_ReturnCount()
    
    'If there are more entries on the list besides this first one, update all corresponding menus now
    If (numEntries > 1) Then
    
        'Start by unloading any/all existing menus
        Dim i As Long
        For i = FormMain.MnuRecentMacros.Count - 1 To 1 Step -1
            Unload FormMain.MnuRecentMacros(i)
        Next i
        
        'Next, load all relevant menus in turn
        For i = 1 To numEntries - 1
            Load FormMain.MnuRecentMacros(i)
        Next i
        
    End If
    
    'Because the "clear list" menu entry sits at the bottom, we must refresh icons when the list is modified
    IconsAndCursors.ResetMenuIcons
    
End Sub

'When loading an entry from file, the setup process is somewhat more streamlined, as we don't need to unload existing
' menu entries.
Private Sub IMRUList_OnItemLoad(ByRef mruManager As pdMRUManager, ByVal Index As Long, ByVal mruCaption As String)
    
    'Note that we do not need to load index 0 - it is always loaded.  Only load menu items greater than this.
    If (Index > 0) Then
        Load FormMain.MnuRecentMacros(Index)
    Else
        FormMain.MnuRecentMacros(Index).Enabled = True
    End If
    
End Sub

Private Sub IMRUList_AfterListLoad(ByRef mruManager As pdMRUManager)
    
    'Make sure the "Clear MRU" option is visible and enabled
    FormMain.MnuClearRecentMacros.Enabled = (mruManager.MRU_ReturnCount > 0)
    
    'Because the "clear list" menu entry sits at the bottom, we must refresh icons when the list is modified
    IconsAndCursors.ResetMenuIcons
    
End Sub

'When the list is empty, leave the first entry of the sub-menu, but disable it
Private Sub IMRUList_OnListIsEmpty()

    FormMain.MnuRecentMacros(0).Enabled = False
    FormMain.MnuClearRecentMacros.Enabled = False
    
    'All icons in this menu need to be manually reset after the list is cleared; the ResetMenuIcons function
    ' will also call the Menus.UpdateSpecialMenu_RecentFiles() function to set all captions properly.
    IconsAndCursors.ResetMenuIcons
    
End Sub

'Unneeded interface methods are collected here
Private Sub IMRUList_BeforeListLoad(ByRef mruManager As pdMRUManager, Optional ByRef Cancel As Boolean = False)

End Sub

Private Sub IMRUList_BeforeListSave(ByRef mruManager As pdMRUManager, Optional ByRef Cancel As Boolean = False)

End Sub

Private Sub IMRUList_AfterListSave(ByRef mruManager As pdMRUManager)

End Sub

Private Sub IMRUList_BeforeFileAdded(ByRef mruManager As pdMRUManager, ByVal newFile As String, Optional ByRef Cancel As Boolean = False)

End Sub

Private Sub IMRUList_SaveThumbnailImage(ByRef filePath As String, ByRef srcImage As pdImage)
    
End Sub

Private Sub IMRUList_OnMaxLimitChanged(ByRef mruManager As pdMRUManager)
    
End Sub

Private Sub IMRUList_EraseEntryAtPosition(ByRef mruManager As pdMRUManager, ByVal Index As Long)

End Sub
