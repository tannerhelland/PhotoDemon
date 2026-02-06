Attribute VB_Name = "Mutex"
'***************************************************************************
'Simplified pdMutex wrapper
'Copyright 2020-2026 by Tanner Helland
'Created: 14/September/20
'Last updated: 14/September/20
'Last update: initial build
'
'This module is just a convenience wrapper for the main pdMutex instance PS uses to check for
' parallel sessions.  Check out the pdMutex class for the interesting bits.
'
'(Note also that this class provides non-mutex workarounds for "unique session detection" to avoid
' VB6 IDE quirks; if you absolutely don't want those, manage your own pdMutex instance.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Const DISABLE_MUTEX_IN_IDE As Boolean = True

Private m_Mutex As pdMutex

Private m_HaveCheckedAlready As Boolean, m_PrevCheckValue As Boolean

'Note: this function only actually checks for multiple instances *once*, by design.
' After that, it returns a cache of the original value.  (Note also that this function
' provides a cheap IDE workaround; the underlying pdMutex object may not exist inside the IDE,
' and multi-instancing checks may not work correctly.)
Public Function IsThisOnlyInstance() As Boolean
    
    'Only check instancing once, at startup time
    If (Not m_HaveCheckedAlready) Then
        
        If (m_Mutex Is Nothing) Then Set m_Mutex = New pdMutex
        
        If (Not OS.IsProgramCompiled) Then
            If (Not DISABLE_MUTEX_IN_IDE) Then
                IsThisOnlyInstance = (Not m_Mutex.DoesMutexAlreadyExist(UserPrefs.GetUniqueAppID(), True))
            Else
                IsThisOnlyInstance = True
            End If
        Else
            IsThisOnlyInstance = (Not m_Mutex.DoesMutexAlreadyExist(UserPrefs.GetUniqueAppID(), True))
        End If
        
        m_HaveCheckedAlready = True
        m_PrevCheckValue = IsThisOnlyInstance
    
    'On subsequent calls, return the initial value we retrieved
    Else
        IsThisOnlyInstance = m_PrevCheckValue
    End If
        
End Function

Public Sub FreeAllMutexes()
    Set m_Mutex = Nothing
End Sub
