VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Simple Mutex class
'Copyright 2020-2026 by Tanner Helland
'Created: 14/September/20
'Last updated: 14/September/20
'Last update: initial build
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateMutexW Lib "kernel32" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As Long) As Long

Private Const ERROR_ALREADY_EXISTS As Long = &HB7&

Private m_hMutex As Long, m_MutexName As String

Private Sub CloseMutex()
    If (m_hMutex <> 0) Then
        If (CloseHandle(m_hMutex) = 0) Then PDDebug.LogAction "WARNING!  pdMutex failed to close its handle; last error was " & Err.LastDllError()
        m_hMutex = 0
        m_MutexName = vbNullString
    End If
End Sub

'Returns TRUE if mutex was created successfully; FALSE if mutex was not created.  If this function failed,
' LastDLLError *is returned automatically* - so don't manually query it!
Friend Function CreateUniqueNamedMutex(ByRef srcMutexName As String, Optional ByRef dstLastErrorIfAny As Long) As Boolean
    
    'Free existing handles before creating a new one
    If (m_hMutex <> 0) Then CloseMutex
    
    m_hMutex = CreateMutexW(0&, 1&, StrPtr(srcMutexName))
    dstLastErrorIfAny = Err.LastDllError
    
    'If this mutex already exists, return FALSE
    If (dstLastErrorIfAny = ERROR_ALREADY_EXISTS) Then
        CreateUniqueNamedMutex = False
        CloseMutex  'Release our copy, but obviously the mutex won't be deleted until *all* handles are released
    
    'This mutex doesn't exist; return TRUE if we received a valid handle
    Else
        CreateUniqueNamedMutex = (m_hMutex <> 0)
        If CreateUniqueNamedMutex Then m_MutexName = srcMutexName
    End If
    
End Function

'Returns TRUE if a mutex with name srcMutexName already exists.
' Optional parameter createIfMissing will leave the mutex open if does not already exist.
Friend Function DoesMutexAlreadyExist(ByRef srcMutexName As String, Optional createIfMissing As Boolean = False) As Boolean
    
    'Try to create the named mutex
    Dim copyOfLastErr As Long
    If CreateUniqueNamedMutex(srcMutexName, copyOfLastErr) Then
        DoesMutexAlreadyExist = False
        If (Not createIfMissing) Then CloseMutex
    
    'We couldn't create the named mutex.  Use lastDllError to return a useful value.
    Else
        DoesMutexAlreadyExist = (copyOfLastErr = ERROR_ALREADY_EXISTS)
    End If
    
End Function

Private Sub Class_Terminate()
    CloseMutex
End Sub
