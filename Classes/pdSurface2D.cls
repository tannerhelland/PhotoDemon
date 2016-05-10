VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSurface2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 2D Rendering Surface class
'Copyright 2012-2016 by Tanner Helland
'Created: 01/September/12
'Last updated: 10/May/16
'Last update: split surface-specific capabilities out from GDI+ and into a safe wrapper
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'The current surface type.  At present, this is immutable; once set, it cannot be migrated
Private m_SurfaceBackend As PD_2D_RENDERING_BACKEND

'The current surface handle.  Will be 0 if the surface is uninitialized.
Private m_SurfaceHandle As Long

'If debug mode is active, classes will report their destruction back to the central Drawing2D module
Private m_DebugMode As Boolean

Public Function CreateSurfaceFromDC(ByVal srcDC As Long, Optional ByVal enableAA As Boolean = True) As Boolean
    
    If (m_SurfaceHandle <> 0) Then ReleaseSurface
    
    Select Case m_SurfaceBackend
        
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(srcDC, enableAA)
            CreateSurfaceFromDC = (m_SurfaceHandle <> 0)
            
    End Select
    
    'When debug mode is active, all surface creations are reported back to the master Drawing2D module
    If (CreateSurfaceFromDC And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceChange m_SurfaceBackend, True
    
End Function

Public Function GetHandle() As Long
    GetHandle = m_SurfaceHandle
End Function

Friend Function GetSurfaceBackend() As PD_2D_RENDERING_BACKEND
    GetSurfaceBackend = m_SurfaceBackend
End Function

Public Function HasSurface() As Boolean
    HasSurface = CBool(m_SurfaceHandle <> 0)
End Function

'You only need to manually initialize a surface if you plan on using a non-default backend.  (Surfaces with the
' default backend are initialized automatically.)
Friend Function InitializeSurface(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = PD2D_DefaultBackend) As Boolean
    m_SurfaceBackend = targetBackend
End Function

Public Function ReleaseSurface() As Boolean
    
    If (m_SurfaceHandle <> 0) Then
    
        Select Case m_SurfaceBackend
        
            Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
                ReleaseSurface = GDI_Plus.ReleaseGDIPlusGraphics(m_SurfaceHandle)
        
        End Select
        
        'After a successful release, we must always reset the class-level handle to match
        If ReleaseSurface Then
            m_SurfaceHandle = 0
            
            'When debug mode is active, all surface creations are reported back to the master Drawing2D module
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceChange m_SurfaceBackend, False
        End If
    
    Else
        ReleaseSurface = True
    End If

End Function

Public Function GetSurfaceAAMode() As Boolean
    Select Case m_SurfaceBackend
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            GetSurfaceAAMode = GDI_Plus.GetGraphicsAntialiasing(m_SurfaceHandle)
    End Select
End Function

Public Function SetSurfaceAAMode(ByVal newMode As Boolean) As Boolean
    Select Case m_SurfaceBackend
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            SetSurfaceAAMode = GDI_Plus.SetGraphicsAntialiasing(m_SurfaceHandle, newMode)
    End Select
End Function

'When debug mode is active, this class will report surface creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Terminate()
    ReleaseSurface
End Sub
