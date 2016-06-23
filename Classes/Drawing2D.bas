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
    
    'As a convenience, gradient brushes can be fully get/set as a whole XML packet (e.g. the XML returned by
    ' the pd2DGradient class).  This overrides all existing gradient settings.
    P2_BrushGradientAllSettings = 8
    
    'Most individual gradient settings can also be get/set individually (with the exception of nodes, which must
    ' be handled as an entire group - this is a limitation of pdGradient, specifically)
    P2_BrushGradientShape = 9
    P2_BrushGradientAngle = 10
    P2_BrushGradientWrapMode = 11
    P2_BrushGradientNodes = 12
    
    'Textures are somewhat problematic because we store them inside a DIB, which is not easily serializable.  Solving this
    ' is TODO; there's always Base-64, obviously, although performance ain't gonna be great.
    P2_BrushTextureWrapMode = 13
    
    [_P2_NumOfBrushSettings] = 14
End Enum

#If False Then
    Private Const P2_BrushMode = 0, P2_BrushColor = 1, P2_BrushOpacity = 2, P2_BrushPatternStyle = 3, P2_BrushPattern1Color = 4, P2_BrushPattern1Opacity = 5, P2_BrushPattern2Color = 6, P2_BrushPattern2Opacity = 7, P2_BrushGradientAllSettings = 8, P2_BrushGradientShape = 9, P2_BrushGradientAngle = 10
    Private Const P2_BrushGradientWrapMode = 11, P2_BrushGradientNodes = 12, P2_BrushTextureWrapMode = 13, P2_NumOfBrushSettings = 14
#End If

'Gradients work a little differently; they expose *some* properties that you can change directly, but things like
' individual gradient points must be operated on through dedicated functions.
Public Enum PD_2D_GRADIENT_SETTINGS
    P2_GradientShape = 0
    P2_GradientAngle = 1
    P2_GradientWrapMode = 2
    P2_GradientNodes = 3
End Enum

#If False Then
    Private Const P2_GradientShape = 0, P2_GradientAngle = 1, P2_GradientWrapMode = 2, P2_GradientNodes = 3
#End If

'Surfaces are somewhat limited at present, but this may change in the future
Public Enum PD_2D_SURFACE_SETTINGS
    P2_SurfaceAntialiasing = 0
    P2_SurfacePixelOffset = 1
    P2_SurfaceRenderingOriginX = 2
    P2_SurfaceRenderingOriginY = 3
    P2_SurfaceBlendUsingSRGBGamma = 4
    [_P2_NumOfSurfaceSettings] = 5
End Enum

#If False Then
    Private Const P2_SurfaceAntialiasing = 0, P2_SurfacePixelOffset = 1, P2_SurfaceRenderingOriginX = 2, P2_SurfaceRenderingOriginY = 3, P2_SurfaceBlendUsingSRGBGamma = 4, P2_NumOfSurfaceSettings = 5
#End If

'The whole point of Drawing2D is to avoid backend-specific parameters.  As such, we necessarily wrap a number of
' GDI+ enums with our own P2-prefixed enums.  This seems redundant (and it is), but this is what makes it possible
' to support backends with different capabilities.
'
'As such, all Drawing2D classes operate on the enums defined in this class.  Where appropriate, they internally
' remap these values to backend-specific ones.

Public Enum PD_2D_Antialiasing
    P2_AA_None = 0&
    P2_AA_HighQuality = 1&
End Enum

#If False Then
    Private Const P2_AA_None = 0&, P2_AA_HighQuality = 1&
#End If

Public Enum PD_2D_BrushMode
    P2_BM_Solid = 0
    P2_BM_Pattern = 1
    P2_BM_Gradient = 2
    P2_BM_Texture = 3
End Enum

#If False Then
    Private Const P2_BM_Solid = 0, P2_BM_Pattern = 1, P2_BM_Gradient = 2, P2_BM_Texture = 3
#End If

Public Enum PD_2D_CombineMode
    P2_CM_Replace = 0
    P2_CM_Intersect = 1
    P2_CM_Union = 2
    P2_CM_Xor = 3
    P2_CM_Exclude = 4
    P2_CM_Complement = 5
End Enum

#If False Then
    Private Const P2_CM_Replace = 0, P2_CM_Intersect = 1, P2_CM_Union = 2, P2_CM_Xor = 3, P2_CM_Exclude = 4, P2_CM_Complement = 5
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

Public Enum PD_2D_FillRule
    P2_FR_OddEven = 0&
    P2_FR_Winding = 1&
End Enum

#If False Then
    Private Const P2_FR_OddEven = 0&, P2_FR_Winding = 1&
#End If

Public Enum PD_2D_GradientShape
    P2_GS_Linear = 0
    P2_GS_Reflection = 1
    P2_GS_Radial = 2
    P2_GS_Rectangle = 3
    P2_GS_Diamond = 4
End Enum

#If False Then
    Private Const P2_GS_Linear = 0, P2_GS_Reflection = 1, P2_GS_Radial = 2, P2_GS_Rectangle = 3, P2_GS_Diamond = 4
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

'Surfaces come in a few different varieties.  Note that some actions may not be available for certain surface types.
Public Enum PD_2D_SurfaceType
    P2_ST_Uninitialized = -1    'The default value of a new surface; the surface is empty, and cannot be painted to
    P2_ST_WrapperOnly = 0       'This surface is just a wrapper around an existing hDC; pdSurface did not create it
    P2_ST_Bitmap = 1            'This surface is a bitmap (raster) surface, created and owned by a pdSurface instance
End Enum

#If False Then
    Private Const P2_ST_WrapperOnly = 0, P2_ST_Bitmap = 1
#End If

Public Enum PD_2D_TransformOrder
    P2_TO_Prepend = 0&
    P2_TO_Append = 1&
End Enum

#If False Then
    Private Const P2_TO_Prepend = 0&, P2_TO_Append = 1&
#End If

Public Enum PD_2D_WrapMode
    P2_WM_Tile = 0
    P2_WM_TileFlipX = 1
    P2_WM_TileFlipY = 2
    P2_WM_TileFlipXY = 3
    P2_WM_Clamp = 4
End Enum

#If False Then
    Private Const P2_WM_Tile = 0, P2_WM_TileFlipX = 1, P2_WM_TileFlipY = 2, P2_WM_TileFlipXY = 3, P2_WM_Clamp = 4
#End If

'Certain structs are immensely helpful when drawing
Public Type RGBQUAD
   Blue As Byte
   Green As Byte
   Red As Byte
   Alpha As Byte
End Type

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

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Public Type SAFEARRAYBOUND
    cElements As Long
    lBound   As Long
End Type

Public Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Public Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements As Long
    lBound   As Long
End Type

'PD's gradient format is straightforward, and it's declared here so functions can easily create their own gradient interfaces.
Public Type GRADIENTPOINT
    PointRGB As Long
    PointOpacity As Single
    PointPosition As Single
End Type

'Many drawing features lean on various geometry functions
Public Const PI As Double = 3.14159265358979
Public Const PI_HALF As Double = 1.5707963267949
Public Const PI_DOUBLE As Double = 6.28318530717958
Public Const PI_DIV_180 As Double = 0.017453292519943

'If GDI+ is initialized successfully, this will be set to TRUE
Private m_GDIPlusAvailable As Boolean

'When debug mode is active, this module will track surface creation+destruction counts.  This is helpful for detecting leaks.
Private m_DebugMode As Boolean

'When debug mode is active, live counts of various drawing objects are tracked on a per-backend basis.  This is crucial for
' leak detection - these numbers should always match the number of active class instances.
Private m_BrushCount_GDIPlus As Long, m_PathCount_GDIPlus As Long, m_PenCount_GDIPlus As Long, m_RegionCount_GDIPlus As Long, m_SurfaceCount_GDIPlus As Long, m_TransformCount_GDIPlus As Long

'Helper color functions for moving individual RGB components between RGB() Longs.  Note that these functions only
' return values in the range [0, 255], but declaring them as integers prevents overflow during intermediary math steps.
Public Function ExtractRed(ByVal srcColor As Long) As Integer
    ExtractRed = srcColor And 255
End Function

Public Function ExtractGreen(ByVal srcColor As Long) As Integer
    ExtractGreen = (srcColor \ 256) And 255
End Function

Public Function ExtractBlue(ByVal srcColor As Long) As Integer
    ExtractBlue = (srcColor \ 65536) And 255
End Function

'Shortcut function for creating a generic painter
Public Function QuickCreatePainter(ByRef dstPainter As pd2DPainter) As Boolean
    If (dstPainter Is Nothing) Then Set dstPainter = New pd2DPainter
    dstPainter.SetDebugMode m_DebugMode
    QuickCreatePainter = True
End Function

'Shortcut function for creating a new rectangular region with the default rendering backend
Public Function QuickCreateRegionRectangle(ByRef dstRegion As pd2DRegion, ByVal rLeft As Single, ByVal rTop As Single, ByVal rWidth As Single, ByVal rHeight As Single) As Boolean
    If (dstRegion Is Nothing) Then Set dstRegion = New pd2DRegion Else dstRegion.ResetAllProperties
    With dstRegion
        .SetDebugMode m_DebugMode
        QuickCreateRegionRectangle = .AddRectangleF(rLeft, rTop, rWidth, rHeight, P2_CM_Replace)
    End With
End Function

'Shortcut function for quickly creating a blank surface with the default rendering backend and default rendering settings
Public Function QuickCreateBlankSurface(ByRef dstSurface As pd2DSurface, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, Optional ByVal surfaceSupportsAlpha As Boolean = True, Optional ByVal enableAntialiasing As Boolean = False, Optional ByVal initialColor As Long = vbWhite, Optional ByVal initialOpacity As Single = 100#) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        .SetDebugMode m_DebugMode
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateBlankSurface = .CreateBlankSurface(surfaceWidth, surfaceHeight, surfaceSupportsAlpha, initialColor, initialOpacity)
    End With
End Function

'Shortcut function for creating a new surface with the default rendering backend and default rendering settings
Public Function QuickCreateSurfaceFromDC(ByRef dstSurface As pd2DSurface, ByVal srcDC As Long, Optional ByVal enableAntialiasing As Boolean = False) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        .SetDebugMode m_DebugMode
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateSurfaceFromDC = .WrapSurfaceAroundDC(srcDC)
    End With
End Function

'Shortcut function for creating a solid brush
Public Function QuickCreateSolidBrush(ByRef dstBrush As pd2DBrush, Optional ByVal brushColor As Long = vbWhite, Optional ByVal brushOpacity As Single = 100#) As Boolean
    If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush Else dstBrush.ResetAllProperties
    With dstBrush
        .SetDebugMode m_DebugMode
        .SetBrushColor brushColor
        .SetBrushOpacity brushOpacity
        QuickCreateSolidBrush = .CreateBrush()
    End With
End Function

'Shortcut function for creating a two-color gradient brush
Public Function QuickCreateTwoColorGradientBrush(ByRef dstBrush As pd2DBrush, ByRef gradientBoundary As RECTF, Optional ByVal firstColor As Long = vbBlack, Optional ByVal secondColor As Long = vbWhite, Optional ByVal firstColorOpacity As Single = 100#, Optional ByVal secondColorOpacity As Single = 100#, Optional ByVal gradientShape As PD_2D_GradientShape = P2_GS_Linear, Optional ByVal gradientAngle As Single = 0#) As Boolean
    
    If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush Else dstBrush.ResetAllProperties
    
    Dim tmpGradient As pd2DGradient
    Set tmpGradient = New pd2DGradient
    With tmpGradient
        .SetGradientShape gradientShape
        .SetGradientAngle gradientAngle
        .CreateTwoPointGradient firstColor, secondColor, firstColorOpacity, secondColorOpacity
    End With
    
    With dstBrush
        .SetDebugMode m_DebugMode
        .SetBrushMode P2_BM_Gradient
        .SetBoundaryRect gradientBoundary
        .SetBrushGradientAllSettings tmpGradient.GetGradientAsString
        QuickCreateTwoColorGradientBrush = .CreateBrush()
    End With
    
End Function

'Shortcut function for creating a solid pen
Public Function QuickCreateSolidPen(ByRef dstPen As pd2DPen, Optional ByVal penWidth As Single = 1#, Optional ByVal penColor As Long = vbWhite, Optional ByVal penOpacity As Single = 100#, Optional ByVal penLineJoin As PD_2D_LineJoin = P2_LJ_Miter, Optional ByVal penLineCap As PD_2D_LineCap = P2_LC_Flat) As Boolean
    If (dstPen Is Nothing) Then Set dstPen = New pd2DPen Else dstPen.ResetAllProperties
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

'Shortcut function for creating two pens for UI purposes.  This function could use a clearer name, but "UI pens" consist
' of a wide, semi-translucent black pen on bottom, and a thin, less-translucent white pen on top.  This combination
' of pens are perfect for drawing on any arbitrary background of any color or pattern, and ensuring the line will still
' be visible.  (This approach is used in modern software instead of the old "invert" pen approach of past decades.)
'
'If the line is currently being hovered or otherwise interacted with, you can set "useHighlightColor" to TRUE.  This will
' return the top pen in the current highlight color (hard-coded at the top of this module) instead of plain white.
'
'By design, pen width is not settable via this function.  The top pen will always be 1.6 pixels wide (a size required
' to bypass GDI+ subpixel flaws between 1 and 2 pixels) while the bottom pen will always be 3.0 pixels wide.
Public Function QuickCreatePairOfUIPens(ByRef dstPenBase As pd2DPen, ByRef dstPenTop As pd2DPen, Optional ByVal useHighlightColor As Boolean = False, Optional ByVal penLineJoin As PD_2D_LineJoin = P2_LJ_Miter, Optional ByVal penLineCap As PD_2D_LineCap = P2_LC_Flat) As Boolean
    
    If (dstPenBase Is Nothing) Then Set dstPenBase = New pd2DPen Else dstPenBase.ResetAllProperties
    If (dstPenTop Is Nothing) Then Set dstPenTop = New pd2DPen Else dstPenTop.ResetAllProperties
    
    With dstPenBase
        .SetDebugMode m_DebugMode
        .SetPenWidth 3#
        .SetPenColor vbBlack
        .SetPenOpacity 75#
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreatePairOfUIPens = .CreatePen()
    End With
    
    With dstPenTop
        .SetDebugMode m_DebugMode
        .SetPenWidth 1.6
        If useHighlightColor Then .SetPenColor RGB(80, 145, 255) Else .SetPenColor vbWhite
        .SetPenOpacity 87.5
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreatePairOfUIPens = CBool(QuickCreatePairOfUIPens And .CreatePen())
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

Public Function GetLibraryDebugMode() As Boolean
    GetLibraryDebugMode = m_DebugMode
End Function

Public Sub SetLibraryDebugMode(ByVal newMode As Boolean)
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
            InternalError "Bad Parameter", "Couldn't start requested backend: backend ID unknown"
    
    End Select

End Function

'Stop a running rendering backend
Public Function StopRenderingEngine(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    
    Select Case targetBackend
            
        Case P2_DefaultBackend, P2_GDIPlusBackend
            StopRenderingEngine = GDI_Plus.GDIP_StopEngine()
            m_GDIPlusAvailable = False
            
        Case Else
            InternalError "Bad Parameter", "Couldn't stop requested backend: backend ID unknown"
    
    End Select
    
End Function

'At present, Drawing2D errors are simply forwarded to the main error handler function at the bottom of this module.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    DEBUG_NotifyExternalError errName, errDescription, ErrNum, "Drawing2d"
End Sub

'DEBUG FUNCTIONS FOLLOW.  These functions should not be called directly.  They are invoked by other pd2D class when m_DebugMode = TRUE.
Public Sub DEBUG_NotifyBrushCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_BrushCount_GDIPlus = m_BrushCount_GDIPlus + 1 Else m_BrushCount_GDIPlus = m_BrushCount_GDIPlus - 1
        Case Else
            InternalError "Bad Parameter", "Brush creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifyPathCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_PathCount_GDIPlus = m_PathCount_GDIPlus + 1 Else m_PathCount_GDIPlus = m_PathCount_GDIPlus - 1
        Case Else
            InternalError "Bad Parameter", "Path creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifyPenCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_PenCount_GDIPlus = m_PenCount_GDIPlus + 1 Else m_PenCount_GDIPlus = m_PenCount_GDIPlus - 1
        Case Else
            InternalError "Bad Parameter", "Pen creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifyRegionCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_RegionCount_GDIPlus = m_RegionCount_GDIPlus + 1 Else m_RegionCount_GDIPlus = m_RegionCount_GDIPlus - 1
        Case Else
            InternalError "Bad Parameter", "Region creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifySurfaceCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus + 1 Else m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus - 1
        Case Else
            InternalError "Bad Parameter", "Surface creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

Public Sub DEBUG_NotifyTransformCountChange(ByVal targetBackend As PD_2D_RENDERING_BACKEND, ByVal objectCreated As Boolean)
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If objectCreated Then m_TransformCount_GDIPlus = m_TransformCount_GDIPlus + 1 Else m_TransformCount_GDIPlus = m_TransformCount_GDIPlus - 1
        Case Else
            InternalError "Bad Parameter", "Transform creation/destruction was not counted: backend ID unknown"
    End Select
End Sub

'In a default build, external pd2D classes relay any internal errors to this function.  You may wish to modify those classes
' to raise their own error events, or perhaps handle their errors internally.  (By default, pd2D does *not* halt on errors.)
Public Sub DEBUG_NotifyExternalError(Optional ByVal errName As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0, Optional ByVal errSource As String = vbNullString)
    If m_DebugMode Then
        If (Len(errSource) = 0) Then errSource = "pd2D"
        Debug.Print "WARNING!  " & errSource & " encountered an error: """ & errName & """ - " & errDescription
        If (ErrNum <> 0) Then Debug.Print "  (If it helps, an error number was also reported: #" & ErrNum & ")"
    End If
End Sub
