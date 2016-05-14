VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DSurface"
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

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Surface settings that support get/set operations
Private m_SurfaceAntialiasing As PD_2D_Antialiasing
Private m_SurfacePixelOffset As PD_2D_PixelOffset

'Get/set individual settings.  Note that these just wrap the generic Get/SetPenProperty functions, below.
Public Function GetSurfaceAntialiasing() As PD_2D_Antialiasing
    GetSurfaceAntialiasing = Me.GetSurfaceProperty(P2_SurfaceAntialiasing)
End Function

Public Function GetSurfacePixelOffset() As PD_2D_PixelOffset
    GetSurfacePixelOffset = Me.GetSurfaceProperty(P2_SurfacePixelOffset)
End Function

Public Sub SetSurfaceAntialiasing(ByVal newSetting As PD_2D_Antialiasing)
    Me.SetSurfaceProperty P2_SurfaceAntialiasing, newSetting
End Sub

Public Sub SetSurfacePixelOffset(ByVal newSetting As PD_2D_PixelOffset)
    Me.SetSurfaceProperty P2_SurfacePixelOffset, newSetting
End Sub

'Get/set individual settings by enum type
Public Function GetSurfaceProperty(ByVal propID As PD_2D_SURFACE_SETTINGS) As Variant

    Select Case propID
    
        Case P2_SurfaceAntialiasing
            GetSurfaceProperty = m_SurfaceAntialiasing
        
        Case P2_SurfacePixelOffset
            GetSurfaceProperty = m_SurfacePixelOffset
    
    End Select

End Function

Public Sub SetSurfaceProperty(ByVal propID As PD_2D_SURFACE_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
    
        Case P2_SurfaceAntialiasing
            m_SurfaceAntialiasing = CLng(newSetting)
        
        Case P2_SurfacePixelOffset
            m_SurfacePixelOffset = CLng(newSetting)
            
    End Select
    
    RelayPropertyChangeToBackend propID
    
End Sub

Public Function WrapSurfaceAroundDC(ByVal srcDC As Long) As Boolean
    
    If (m_SurfaceHandle <> 0) Then ReleaseSurface
    
    'Call the backend-specific surface creation function
    Select Case m_SurfaceBackend
        
        Case P2_DefaultBackend, P2_GDIPlusBackend
            m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(srcDC, IIf(m_SurfaceAntialiasing >= P2_AA_Grayscale, GP_SM_AntiAlias, GP_SM_None), IIf(m_SurfacePixelOffset >= P2_PO_Normal, GP_POM_Half, GP_POM_None))
            WrapSurfaceAroundDC = (m_SurfaceHandle <> 0)
            
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    If (WrapSurfaceAroundDC And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, True
    
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

'You only need to manually initialize a surface backend if you plan on using a non-default backend.  (Surfaces with the
' default backend are initialized automatically.)
Friend Function InitializeSurfaceBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_SurfaceBackend = targetBackend
    InitializeSurfaceBackend = True
End Function

Public Function ReleaseSurface() As Boolean
    
    If (m_SurfaceHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_SurfaceBackend
        
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleaseSurface = GDI_Plus.ReleaseGDIPlusGraphics(m_SurfaceHandle)
        
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseSurface Then
            m_SurfaceHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, False
        End If
    
    Else
        ReleaseSurface = True
    End If

End Function

Private Sub RelayPropertyChangeToBackend(ByVal propID As PD_2D_SURFACE_SETTINGS)
    
    'Property changes only need to be relayed if we've already created a handle
    If (m_SurfaceHandle <> 0) Then
    
        'Relays obviously vary by the active backend!
        Select Case m_SurfaceBackend
            
            Case P2_DefaultBackend, P2_GDIPlusBackend
                
                Select Case propID
                    Case P2_SurfaceAntialiasing
                        If (m_SurfaceAntialiasing >= P2_AA_Grayscale) Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_SM_AntiAlias
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_SM_None
                        End If
                        
                    Case P2_SurfacePixelOffset
                        If (m_SurfacePixelOffset >= P2_PO_Half) Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_POM_Half
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_POM_None
                        End If
                        
                End Select
                
            'In the future, other backends can be handled here...
            Case Else
                
        End Select
    
    End If
        
End Sub

Public Sub ResetAllProperties()
    Me.SetSurfaceProperty P2_SurfaceAntialiasing, P2_AA_None
    Me.SetSurfaceProperty P2_SurfacePixelOffset, P2_PO_Normal
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    Me.ResetAllProperties
End Sub

Private Sub Class_Terminate()
    ReleaseSurface
End Sub
