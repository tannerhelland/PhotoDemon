Attribute VB_Name = "Drawing2D"
'***************************************************************************
'High-Performance 2D Rendering Interface
'Copyright 2012-2016 by Tanner Helland
'Created: 1/September/12
'Last updated: 09/May/16
'Last update: start migrating various rendering bits out of GDI+ and into this generic renderer.
'
'In 2015-2016, I slowly migrated PhotoDemon to its own UI toolkit.  The new toolkit performs a ton of 2D rendering tasks,
' so it was finally time to migrate PD's hoary old GDI+ interface to a more modern solution.
'
'This module provides a renderer-agnostic solution for various 2D drawing tasks.  At present, it leans only on GDI+,
' but I have tried to design it so that other backends could be used without much trouble.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_2D_RENDERING_BACKEND
    PD2D_DefaultBackend = 0
    PD2D_GDIPlusBackend = 1
End Enum

#If False Then
    Private Const PD2D_DefaultBackend = 0, PD2D_GDIPlusBackend = 1
#End If

'If GDI+ is initialized successfully, this will be set to TRUE
Private m_GDIPlusAvailable As Boolean

'When debug mode is active, this module will track surface creation+destruction counts.  This is helpful for detecting leaks.
Private m_DebugMode As Boolean

'When debug mode is active, live surface counts are tracked for each backend
Private m_SurfaceCount_GDIPlus As Long

'Create a new surface using the default rendering backend
Public Function CreateSurfaceFromDC(ByRef dstSurface As pdSurface2D, ByVal srcDC As Long, Optional ByVal enableAA As Boolean = True) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pdSurface2D
    dstSurface.SetDebugMode m_DebugMode
    CreateSurfaceFromDC = dstSurface.CreateSurfaceFromDC(srcDC, enableAA)
End Function

Public Function IsRenderingEngineActive(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = PD2D_DefaultBackend) As Boolean
    Select Case targetBackend
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            IsRenderingEngineActive = m_GDIPlusAvailable
        Case Else
            IsRenderingEngineActive = False
    End Select
End Function

Public Sub SetDrawing2DDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

'Start a new rendering backend
Public Function StartRenderingBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = PD2D_DefaultBackend) As Boolean

    Select Case targetBackend
            
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            #If DEBUGMODE = 1 Then
                StartRenderingBackend = GDI_Plus.GDIP_StartEngine(True)
            #Else
                StartRenderingBackend = GDI_Plus.GDIP_StartEngine(False)
            #End If
            
            m_GDIPlusAvailable = StartRenderingBackend
            
        Case Else
            InternalRenderingError "Bad Parameter", "Couldn't start requested backend: backend ID unknown"
    
    End Select

End Function

'Stop a running rendering backend
Public Function StopRenderingEngine(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = PD2D_DefaultBackend) As Boolean
    
    Select Case targetBackend
            
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            StopRenderingEngine = GDI_Plus.GDIP_StopEngine()
            m_GDIPlusAvailable = False
            
        Case Else
            InternalRenderingError "Bad Parameter", "Couldn't stop requested backend: backend ID unknown"
    
    End Select
    
End Function

Private Sub InternalRenderingError(Optional ByRef ErrName As String = vbNullString, Optional ByRef ErrDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  Drawing2D encountered an error: """ & ErrName & """ - " & ErrDescription
    #End If

End Sub

'DEBUG FUNCTIONS FOLLOW
Public Sub DEBUG_NotifySurfaceChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal surfaceCreated As Boolean)
    
    Select Case targetBackend
            
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            If surfaceCreated Then m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus + 1 Else m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus - 1
            
        Case Else
            InternalRenderingError "Bad Parameter", "Surface creation/destruction was not counted: backend ID unknown"
    
    End Select
    
End Sub
