VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Region Class
'Copyright 2016-2016 by Tanner Helland
'Created: 18/June/16 (but assembled from many bits written earlier)
'Last updated: 18/June/16
'Last update: start migrating various region bits into this dedicated class
'
'This class manages a single region instance.  Unlike other classes, a default (infinite) region is constructed
' by default, when the class is initialized.  This greatly simplifies subsequent interactions, but it may cause
' some confusion when debugging - so adjust any leak-checking expectations accordingly!
'
'At present, it is primarily based on the capabilities of GDI+.  This may change going forward, but because GDI+
' provides a nice baseline feature set, that's where we started.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This class is not yet capable of serializing itself to/from XML strings, but it may be possible in the future...
Private cSerialize As pdParamXML

'The backend used to manage this region.  At present, this is immutable; once set, it cannot be migrated.
Private m_RegionBackend As PD_2D_RENDERING_BACKEND

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Once a region has been created, this handle value will be non-zero
Private m_RegionHandle As Long

'Add your own error-handling behavior here, as desired
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  pd2DRegion encountered an error: """ & errName & """ - " & errDescription
        If (ErrNum <> 0) Then pdDebug.LogAction "(Also, an error number was reported: " & ErrNum & ")"
    #End If
End Sub

Public Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_RegionHandle = 0)) Then
        If CreateRegion() Then GetHandle = m_RegionHandle Else GetHandle = 0
    Else
        GetHandle = m_RegionHandle
    End If
End Function

Public Function GetRegionBackend() As PD_2D_RENDERING_BACKEND
    GetRegionBackend = m_RegionBackend
End Function

Public Function HasRegion() As Boolean
    HasRegion = CBool(m_RegionHandle <> 0)
End Function

'You only need to manually initialize a backend if you plan on using a NON-DEFAULT backend.
' (Objects with the default backend are initialized automatically.)
Friend Function InitializeRegionBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_RegionBackend = targetBackend
    InitializeRegionBackend = True
End Function

'Create an actual region handle using the current backend and the current region settings.
' NOTE: the caller doesn't *need* to call this directly.  If GetRegionHandle is called and the region doesn't
'       yet exist, it will be auto-created.
Friend Function CreateRegion() As Boolean

    If (m_RegionHandle <> 0) Then ReleaseRegion
    
    'Call the backend-specific creation function
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            m_RegionHandle = GDI_Plus.GetGDIPlusRegionHandle()
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    CreateRegion = CBool(m_RegionHandle <> 0)
    If (CreateRegion And m_DebugMode) Then Drawing2D.DEBUG_NotifyRegionCountChange m_RegionBackend, True
    
End Function

Public Function ReleaseRegion() As Boolean
    
    If (m_RegionHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_RegionBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleaseRegion = GDI_Plus.ReleaseGDIPlusRegion(m_RegionHandle)
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseRegion Then
            m_RegionHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyRegionCountChange m_RegionBackend, False
        End If
    
    Else
        ReleaseRegion = True
    End If
    
End Function

Public Sub ResetAllProperties()
    Me.ReleaseRegion
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    Me.ResetAllProperties
End Sub

Private Sub Class_Terminate()
    Me.ReleaseRegion
End Sub
