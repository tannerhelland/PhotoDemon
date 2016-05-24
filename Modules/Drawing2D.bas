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

'In the future, this module may support multiple different rendering backends.  At present, however, only GDI+ is used.
Public Enum PD_2D_RENDERING_BACKEND
    P2_DefaultBackend = 0
    P2_GDIPlusBackend = 1
End Enum

#If False Then
    Private Const P2_DefaultBackend = 0, P2_GDIPlusBackend = 1
#End If

'To simplify property setting across backends, I use generic enums instead of backend-specific descriptors.
' There are trade-offs with this approach, but I like it because it makes it possible to enumerate object properties.
Public Enum PD_2D_PEN_SETTINGS
    P2_PenStyle = 0
    P2_PenColor = 1
    P2_PenOpacity = 2
    P2_PenWidth = 3
    P2_PenLineJoin = 4
    P2_PenLineCap = 5     'LineCap is a convenience property that sets StartCap, EndCap, and DashCap all at once
    P2_PenDashCap = 6
    P2_PenMiterLimit = 7
    P2_PenAlignment = 8
    P2_PenStartCap = 9
    P2_PenEndCap = 10
    [_P2_NumOfPenSettings] = 11
End Enum

#If False Then
    Private Const P2_PenStyle = 0, P2_PenColor = 1, P2_PenOpacity = 2, P2_PenWidth = 3, P2_PenLineJoin = 4, P2_PenLineCap = 5, P2_PenDashCap = 6, P2_PenMiterLimit = 7, P2_PenAlignment = 8, P2_PenStartCap = 9, P2_PenEndCap = 10, P2_NumOfPenSettings = 11
#End If

'Brushes support a *lot* of internal settings.
Public Enum PD_2D_BRUSH_SETTINGS
    P2_BrushMode = 0
    P2_BrushColor = 1
    P2_BrushOpacity = 2
    P2_BrushPatternStyle = 3
    P2_BrushPattern1Color = 4
    P2_BrushPattern1Opacity = 5
    P2_BrushPattern2Color = 6
    P2_BrushPattern2Opacity = 7
    
    'Note that individual gradient values cannot be set/read.  Gradients are only supported as a complete gradient
    ' XML packet, as supplied by the pdGradient class.
    P2_BrushGradientXML = 8
    [_P2_NumOfBrushSettings] = 9
End Enum

#If False Then
    Const P2_BrushMode = 0, P2_BrushColor = 1, P2_BrushOpacity = 2, P2_BrushPatternStyle = 3, P2_BrushPattern1Color = 4, P2_BrushPattern1Opacity = 5, P2_BrushPattern2Color = 6, P2_BrushPattern2Opacity = 7, P2_BrushGradientXML = 8, P2_NumOfBrushSettings = 9
#End If

'Surfaces are somewhat limited at present, but this may change in the future
Public Enum PD_2D_SURFACE_SETTINGS
    P2_SurfaceAntialiasing = 0
    P2_SurfacePixelOffset = 1
    [_P2_NumOfSurfaceSettings] = 2
End Enum

#If False Then
    Private Const P2_SurfaceAntialiasing = 0, P2_SurfacePixelOffset = 1, P2_NumOfSurfaceSettings = 2
#End If

'The whole point of Drawing2D is to avoid backend-specific parameters.  As such, we necessarily wrap a number of
' GDI+ enums with our own P2-prefixed enums.  This seems redundant (and it is), but this is exactly what makes it
' possible to support future backends that offer different capabilities.
' (NOTE: individual PD2D classes note which of these enums map directly to matching GDI+ enums.)

Public Enum PD_2D_BrushMode
    P2_BM_Solid = 0
    P2_BM_Pattern = 1
    P2_BM_Gradient = 2
    P2_BM_Texture = 3
End Enum

#If False Then
    Private Const P2_BM_Solid = 0, P2_BM_Pattern = 1, P2_BM_Gradient = 2, P2_BM_Texture = 3
#End If

Public Enum PD_2D_DashCap
    P2_DC_Flat = 0
    P2_DC_Square = 1        'NOTE: GDI+ does not support square dash caps - only flat ones - so square simply remaps to flat
    P2_DC_Round = 2
    P2_DC_Triangle = 3
End Enum

#If False Then
    Private Const P2_DC_Flat = 0, P2_DC_Square = 0, P2_DC_Round = 2, P2_DC_Triangle = 3
#End If

Public Enum PD_2D_DashStyle
    P2_DS_Solid = 0&
    P2_DS_Dash = 1&
    P2_DS_Dot = 2&
    P2_DS_DashDot = 3&
    P2_DS_DashDotDot = 4&
    P2_DS_Custom = 5&
End Enum

#If False Then
    Private Const P2_DS_Solid = 0&, P2_DS_Dash = 1&, P2_DS_Dot = 2&, P2_DS_DashDot = 3&, P2_DS_DashDotDot = 4&, P2_DS_Custom = 5&
#End If

Public Enum PD_2D_LineCap
    P2_LC_Flat = 0&
    P2_LC_Square = 1&
    P2_LC_Round = 2&
    P2_LC_Triangle = 3&
    P2_LC_FlatAnchor = &H10
    P2_LC_SquareAnchor = &H11
    P2_LC_RoundAnchor = &H12
    P2_LC_DiamondAnchor = &H13
    P2_LC_ArrowAnchor = &H14
    P2_LC_Custom = &HFF
End Enum

#If False Then
    Private Const P2_LC_Flat = 0, P2_LC_Square = 1, P2_LC_Round = 2, P2_LC_Triangle = 3, P2_LC_FlatAnchor = &H10, P2_LC_SquareAnchor = &H11, P2_LC_RoundAnchor = &H12, P2_LC_DiamondAnchor = &H13, P2_LC_ArrowAnchor = &H14, P2_LC_Custom = &HFF
#End If

Public Enum PD_2D_LineJoin
    P2_LJ_Miter = 0&
    P2_LJ_Bevel = 1&
    P2_LJ_Round = 2&
End Enum

#If False Then
    Private Const P2_LJ_Miter = 0&, P2_LJ_Bevel = 1&, P2_LJ_Round = 2&
#End If

Public Enum PD_2D_PatternStyle
    P2_PS_Horizontal = 0
    P2_PS_Vertical = 1
    P2_PS_ForwardDiagonal = 2
    P2_PS_BackwardDiagonal = 3
    P2_PS_Cross = 4
    P2_PS_DiagonalCross = 5
    P2_PS_05Percent = 6
    P2_PS_10Percent = 7
    P2_PS_20Percent = 8
    P2_PS_25Percent = 9
    P2_PS_30Percent = 10
    P2_PS_40Percent = 11
    P2_PS_50Percent = 12
    P2_PS_60Percent = 13
    P2_PS_70Percent = 14
    P2_PS_75Percent = 15
    P2_PS_80Percent = 16
    P2_PS_90Percent = 17
    P2_PS_LightDownwardDiagonal = 18
    P2_PS_LightUpwardDiagonal = 19
    P2_PS_DarkDownwardDiagonal = 20
    P2_PS_DarkUpwardDiagonal = 21
    P2_PS_WideDownwardDiagonal = 22
    P2_PS_WideUpwardDiagonal = 23
    P2_PS_LightVertical = 24
    P2_PS_LightHorizontal = 25
    P2_PS_NarrowVertical = 26
    P2_PS_NarrowHorizontal = 27
    P2_PS_DarkVertical = 28
    P2_PS_DarkHorizontal = 29
    P2_PS_DashedDownwardDiagonal = 30
    P2_PS_DashedUpwardDiagonal = 31
    P2_PS_DashedHorizontal = 32
    P2_PS_DashedVertical = 33
    P2_PS_SmallConfetti = 34
    P2_PS_LargeConfetti = 35
    P2_PS_ZigZag = 36
    P2_PS_Wave = 37
    P2_PS_DiagonalBrick = 38
    P2_PS_HorizontalBrick = 39
    P2_PS_Weave = 40
    P2_PS_Plaid = 41
    P2_PS_Divot = 42
    P2_PS_DottedGrid = 43
    P2_PS_DottedDiamond = 44
    P2_PS_Shingle = 45
    P2_PS_Trellis = 46
    P2_PS_Sphere = 47
    P2_PS_SmallGrid = 48
    P2_PS_SmallCheckerBoard = 49
    P2_PS_LargeCheckerBoard = 50
    P2_PS_OutlinedDiamond = 51
    P2_PS_SolidDiamond = 52
End Enum

#If False Then
    Private Const P2_PS_Horizontal = 0, P2_PS_Vertical = 1, P2_PS_ForwardDiagonal = 2, P2_PS_BackwardDiagonal = 3, P2_PS_Cross = 4, P2_PS_DiagonalCross = 5, P2_PS_05Percent = 6, P2_PS_10Percent = 7, P2_PS_20Percent = 8, P2_PS_25Percent = 9, P2_PS_30Percent = 10, P2_PS_40Percent = 11, P2_PS_50Percent = 12, P2_PS_60Percent = 13, P2_PS_70Percent = 14, P2_PS_75Percent = 15, P2_PS_80Percent = 16, P2_PS_90Percent = 17, P2_PS_LightDownwardDiagonal = 18, P2_PS_LightUpwardDiagonal = 19, P2_PS_DarkDownwardDiagonal = 20, P2_PS_DarkUpwardDiagonal = 21, P2_PS_WideDownwardDiagonal = 22, P2_PS_WideUpwardDiagonal = 23, P2_PS_LightVertical = 24, P2_PS_LightHorizontal = 25
    Private Const P2_PS_NarrowVertical = 26, P2_PS_NarrowHorizontal = 27, P2_PS_DarkVertical = 28, P2_PS_DarkHorizontal = 29, P2_PS_DashedDownwardDiagonal = 30, P2_PS_DashedUpwardDiagonal = 31, P2_PS_DashedHorizontal = 32, P2_PS_DashedVertical = 33, P2_PS_SmallConfetti = 34, P2_PS_LargeConfetti = 35, P2_PS_ZigZag = 36, P2_PS_Wave = 37, P2_PS_DiagonalBrick = 38, P2_PS_HorizontalBrick = 39, P2_PS_Weave = 40, P2_PS_Plaid = 41, P2_PS_Divot = 42, P2_PS_DottedGrid = 43, P2_PS_DottedDiamond = 44, P2_PS_Shingle = 45, P2_PS_Trellis = 46, P2_PS_Sphere = 47, P2_PS_SmallGrid = 48, P2_PS_SmallCheckerBoard = 49, P2_PS_LargeCheckerBoard = 50
    Private Const P2_PS_OutlinedDiamond = 51, P2_PS_SolidDiamond = 52
#End If

Public Enum PD_2D_PenAlignment
    P2_PA_Center = 0&
    P2_PA_Inset = 1&
End Enum

#If False Then
    Private Const P2_PA_Center = 0&, P2_PA_Inset = 1&
#End If

Public Enum PD_2D_PixelOffset
    P2_PO_Normal = 0
    P2_PO_Half = 1
End Enum

#If False Then
    Private Const P2_PO_Normal = 0, P2_PO_Half = 1
#End If

Public Enum PD_2D_Antialiasing
    P2_AA_None = 0&
    P2_AA_Grayscale = 1&
End Enum

#If False Then
    Private Const P2_AA_None = 0&, P2_AA_Grayscale = 1&
#End If

'Certain structs are immensely helpful when drawing
Public Type POINTFLOAT
   x As Single
   y As Single
End Type

Public Type RECTL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RECTF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

'If GDI+ is initialized successfully, this will be set to TRUE
Private m_GDIPlusAvailable As Boolean

'When debug mode is active, this module will track surface creation+destruction counts.  This is helpful for detecting leaks.
Private m_DebugMode As Boolean

'When debug mode is active, live counts of various drawing objects are tracked on a per-backend basis
Private m_SurfaceCount_GDIPlus As Long, m_PenCount_GDIPlus As Long, m_BrushCount_GDIPlus As Long

'Shortcut function for creating a generic painter
Public Function QuickCreatePainter(ByRef dstPainter As pd2DPainter) As Boolean
    If (dstPainter Is Nothing) Then Set dstPainter = New pd2DPainter
    dstPainter.SetDebugMode m_DebugMode
    QuickCreatePainter = True
End Function

'Shortcut function for creating a new surface with the default rendering backend and default rendering settings
Public Function QuickCreateSurfaceFromDC(ByRef dstSurface As pd2DSurface, ByVal srcDC As Long, Optional ByVal enableAntialiasing As Boolean = False) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface
    With dstSurface
        .SetDebugMode m_DebugMode
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_Grayscale Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateSurfaceFromDC = .WrapSurfaceAroundDC(srcDC)
    End With
End Function

'Shortcut function for creating a solid brush
Public Function QuickCreateSolidBrush(ByRef dstBrush As pd2DBrush, Optional ByVal brushColor As Long = vbWhite, Optional ByVal brushOpacity As Single = 100#) As Boolean
    If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush
    With dstBrush
        .SetDebugMode m_DebugMode
        .SetBrushColor brushColor
        .SetBrushOpacity brushOpacity
        QuickCreateSolidBrush = .CreateBrush()
    End With
End Function

'Shortcut function for creating a solid pen
Public Function QuickCreateSolidPen(ByRef dstPen As pd2DPen, Optional ByVal penWidth As Single = 1#, Optional ByVal penColor As Long = vbWhite, Optional ByVal penOpacity As Single = 100#, Optional ByVal penLineJoin As PD_2D_LineJoin = P2_LJ_Miter, Optional ByVal penLineCap As PD_2D_LineCap = P2_LC_Flat) As Boolean
    If (dstPen Is Nothing) Then Set dstPen = New pd2DPen
    With dstPen
        .SetDebugMode m_DebugMode
        .SetPenWidth penWidth
        .SetPenColor penColor
        .SetPenOpacity penOpacity
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreateSolidPen = .CreatePen()
    End With
End Function

Public Function IsRenderingEngineActive(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            IsRenderingEngineActive = m_GDIPlusAvailable
        Case Else
            IsRenderingEngineActive = False
    End Select
End Function

Public Sub SetDrawing2DDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

'Start a new rendering backend
Public Function StartRenderingBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean

    Select Case targetBackend
            
        Case P2_DefaultBackend, P2_GDIPlusBackend
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
Public Function StopRenderingEngine(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    
    Select Case targetBackend
            
        Case P2_DefaultBackend, P2_GDIPlusBackend
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
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_BrushCount_GDIPlus = m_BrushCount_GDIPlus + 1 Else m_BrushCount_GDIPlus = m_BrushCount_GDIPlus - 1
        Case Else
            InternalRenderingError "Bad Parameter", "Brush creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifyPenCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_PenCount_GDIPlus = m_PenCount_GDIPlus + 1 Else m_PenCount_GDIPlus = m_PenCount_GDIPlus - 1
        Case Else
            InternalRenderingError "Bad Parameter", "Pen creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifySurfaceCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus + 1 Else m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus - 1
        Case Else
            InternalRenderingError "Bad Parameter", "Surface creation/destruction was not counted: backend ID unknown"
    End Select
End Sub
