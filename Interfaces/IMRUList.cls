VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMRUList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'IMRUList : Interface for Recent File (MRU) lists
'Copyright 2005-2015 by Raj Chaudhuri and Tanner Helland
'Created: 16/February/15
'Last updated: 17/February/15
'Last updated by: Raj
'Last update: GetThumbnailPath returns a string
'Dependencies: - pdMRUManager (helper class that manages the XML storage of MRU entries).  Note that pdMRUManager has several
'                 dependencies of its own; those are not listed here, for brevity's sake.
'              - pdImage (for lists where thumbnail images are relevant)
'
'This interface provides properties and methods to deal with the non-generic parts of MRU maintainence, such as the path to
' the XML file where the MRU would be saved, and points in time when the UI needs to be updated. See pdMRURecentFiles.cls
' for a sample implementation.
'
'Many thanks to Raj Chaudhuri for contributing this great feature to PhotoDemon.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Property Get XMLFilename() As String

End Property

Public Property Get SupportsThumbnails() As Boolean

End Property

Public Function GetThumbnailPath(ByRef mruManager As pdMRUManager, ByVal Index As Long) As String

End Function

Public Sub SaveThumbnailImage(ByRef filePath As String, ByRef srcImage As pdImage)

End Sub

Public Sub BeforeListLoad(ByRef mruManager As pdMRUManager, Optional ByRef Cancel As Boolean = False)

End Sub

Public Sub AfterListLoad(ByRef mruManager As pdMRUManager)

End Sub

Public Sub BeforeListSave(ByRef mruManager As pdMRUManager, Optional ByRef Cancel As Boolean = False)

End Sub

Public Sub AfterListSave(ByRef mruManager As pdMRUManager)

End Sub

Public Sub BeforeFileAdded(ByRef mruManager As pdMRUManager, ByVal newFile As String, Optional ByRef Cancel As Boolean = False)

End Sub

Public Sub AfterFileAdded(ByRef mruManager As pdMRUManager, ByVal newFilePath As String, ByVal mruCaption As String)

End Sub

Public Sub AfterListCleared()

End Sub

Public Sub OnItemLoad(ByRef mruManager As pdMRUManager, ByVal Index As Long, ByVal mruCaption As String)

End Sub

Public Sub OnListIsEmpty()

End Sub

Public Sub OnMaxLimitChanged(ByRef mruManager As pdMRUManager)

End Sub

Public Sub EraseEntryAtPosition(ByRef mruManager As pdMRUManager, ByVal Index As Long)

End Sub

