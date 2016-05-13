Attribute VB_Name = "Drawing2D"
'***************************************************************************
'High-Performance Backend-Agnostic 2D Rendering Interface
'Copyright 2012-2016 by Tanner Helland
'Created: 1/September/12
'Last updated: 11/May/16
'Last update: continue migrating various rendering bits out of GDI+ and into this generic renderer.
'
'In 2015-2016, I slowly migrated PhotoDemon to its own UI toolkit.  The new toolkit performs a ton of 2D rendering tasks,
' so it was finally time to migrate PD's hoary old GDI+ interface to a more modern solution.
'
'This module provides a renderer-agnostic solution for various 2D drawing tasks.  At present, it leans only on GDI+,
' but I have tried to design it so that other backends can be supported without much trouble.
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

'To simplify property setting across backends, we use generic enums instead of backend-specific descriptors.
' There are trade-offs with this approach, but I like it because it lets me use single Get/Set functions for
' all of a drawing object's properties (at the cost of not having friendly enums right there in the function
' parameters).  In the future, it would be nice to have both options available, but since dedicated Get/Set
' functions for each property is tedious work, I'm postponing it for now.
Public Enum PD_2D_PEN_SETTINGS
    PD2D_PenStyle = 0
    PD2D_PenColor = 1
    PD2D_PenOpacity = 2
    PD2D_PenWidth = 3
    PD2D_PenLineJoin = 4
    PD2D_PenLineCap = 5     'LineCap is a convenience property that sets StartCap, EndCap, and DashCap all at once
    PD2D_PenDashCap = 6
    PD2D_PenMiterLimit = 7
    PD2D_PenAlignment = 8
    PD2D_PenStartCap = 9
    PD2D_PenEndCap = 10
    [_PD2D_NumOfPenSettings] = 11
End Enum

#If False Then
    Private Const PD2D_PenStyle = 0, PD2D_PenColor = 1, PD2D_PenOpacity = 2, PD2D_PenWidth = 3, PD2D_PenLineJoin = 4, PD2D_PenLineCap = 5, PD2D_PenDashCap = 6, PD2D_PenMiterLimit = 7, PD2D_PenAlignment = 8, PD2D_PenStartCap = 9, PD2D_PenEndCap = 10, PD2D_NumOfPenSettings = 11
#End If

'Brushes support a *lot* of internal settings.
Public Enum PD_2D_BRUSH_SETTINGS
    PD2D_BrushMode = 0
    PD2D_BrushColor = 1
    PD2D_BrushOpacity = 2
    PD2D_BrushPatternStyle = 3
    PD2D_BrushPattern1Color = 4
    PD2D_BrushPattern1Opacity = 5
    PD2D_BrushPattern2Color = 6
    PD2D_BrushPattern2Opacity = 7
    
    'Note that individual gradient values cannot be set/read.  Gradients are only supported as a complete gradient
    ' XML packet, as supplied by the pdGradient class.
    PD2D_BrushGradientXML = 8
    [_PD2D_NumOfBrushSettings] = 9
End Enum

#If False Then
    Const PD2D_BrushMode = 0, PD2D_BrushColor = 1, PD2D_BrushOpacity = 2, PD2D_BrushPatternStyle = 3, PD2D_BrushPattern1Color = 4, PD2D_BrushPattern1Opacity = 5, PD2D_BrushPattern2Color = 6, PD2D_BrushPattern2Opacity = 7, PD2D_BrushGradientXML = 8, PD2D_NumOfBrushSettings = 9
#End If

Public Enum PD_2D_SURFACE_SETTINGS
    PD2D_SurfaceAntialiasing = 0
    PD2D_SurfacePixelOffset = 1
    [_PD2D_NumOfSurfaceSettings] = 2
End Enum

#If False Then
    Private Const PD2D_SurfaceAntialiasing = 0, PD2D_SurfacePixelOffset = 1, PD2D_NumOfSurfaceSettings = 2
#End If

'If GDI+ is initialized successfully, this will be set to TRUE
Private m_GDIPlusAvailable As Boolean

'When debug mode is active, this module will track surface creation+destruction counts.  This is helpful for detecting leaks.
Private m_DebugMode As Boolean

'When debug mode is active, live counts of various drawing objects are tracked on a per-backend basis
Private m_SurfaceCount_GDIPlus As Long, m_PenCount_GDIPlus As Long, m_BrushCount_GDIPlus As Long

'Shortcut function for creating a new surface with the default rendering backend and default rendering settings
Public Function QuickCreateSurfaceFromDC(ByRef dstSurface As pd2DSurface, ByVal srcDC As Long) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface
    With dstSurface
        .SetDebugMode m_DebugMode
        QuickCreateSurfaceFromDC = .WrapSurfaceAroundDC(srcDC)
    End With
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

Private Sub InternalRenderingError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  Drawing2D encountered an error: """ & errName & """ - " & errDescription
    #End If
End Sub

'DEBUG FUNCTIONS FOLLOW.  These functions should not be called directly.  They are invoked by other pd2D class when m_DebugMode = TRUE.
Public Sub DEBUG_NotifyBrushCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            If objectCreated Then m_BrushCount_GDIPlus = m_BrushCount_GDIPlus + 1 Else m_BrushCount_GDIPlus = m_BrushCount_GDIPlus - 1
        Case Else
            InternalRenderingError "Bad Parameter", "Brush creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifyPenCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            If objectCreated Then m_PenCount_GDIPlus = m_PenCount_GDIPlus + 1 Else m_PenCount_GDIPlus = m_PenCount_GDIPlus - 1
        Case Else
            InternalRenderingError "Bad Parameter", "Pen creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifySurfaceCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            If objectCreated Then m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus + 1 Else m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus - 1
        Case Else
            InternalRenderingError "Bad Parameter", "Surface creation/destruction was not counted: backend ID unknown"
    End Select
End Sub
