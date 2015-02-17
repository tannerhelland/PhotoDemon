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
Option Explicit

Public Property Get XMLFileName() As String

End Property

Public Property Get SupportsThumbNails() As Boolean

End Property

Public Function GetThumbNailPath(manager As pdMRUManager, Index As Long)

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

