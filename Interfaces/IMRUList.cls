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
'IMRUList : Interface representing an MRU list
'Copyright 2005-2015 by Tanner Helland
'Created: 16/February/15
'Last updated: 17/February/15
'Last updated by: Raj
'Last update: GetThumbnailPath returns a string
'
'This interface provides properties and methods to deal with the non-generic
'  parts of MRU maintainence, such as the path to the XML file where the MRU
'  would be saved, and points in time when the UI needs to be updated. See
'  pdMRURecentFiles.cls for a sample implementation.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Property Get XMLFileName() As String

End Property

Public Property Get SupportsThumbNails() As Boolean

End Property

Public Function GetThumbNailPath(manager As pdMRUManager, Index As Long) As String

End Function

Public Sub SaveThumbNail(filePath As String, srcImage As pdImage)

End Sub

Public Sub AfterListCleared()

End Sub

Public Sub OnListIsEmpty()

End Sub

Public Sub BeforeListLoad(manager As pdMRUManager, Optional cancel As Boolean = False)

End Sub

Public Sub OnItemLoad(manager As pdMRUManager, Index As Long, mruCaption As String)

End Sub

Public Sub AfterListLoaded(manager As pdMRUManager)

End Sub

Public Sub OnMaxLimitChanged(manager As pdMRUManager)

End Sub

Public Sub BeforeListSave(manager As pdMRUManager, Optional cancel As Boolean = False)

End Sub

Public Sub AfterListSaved(manager As pdMRUManager)

End Sub

Public Sub BeforeNewFileAdded(manager As pdMRUManager, newFile As String, Optional cancel As Boolean = False)

End Sub

Public Sub TrimListAtPosition(manager As pdMRUManager, Index As Long)

End Sub

Public Sub AfterNewFileAdded(manager As pdMRUManager, newFile As String, mruCaption As String)

End Sub

