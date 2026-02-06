Attribute VB_Name = "Drawing2D"
'***************************************************************************
'High-Performance 2D Rendering Interface
'Copyright 2012-2026 by Tanner Helland
'Created: 1/September/12
'Last updated: 26/February/20
'Last update: new helper functions for safer XML serialization of enums
'
'In 2015-2019, I slowly migrated PhotoDemon to its own UI toolkit.  The new toolkit performs a ton
' of 2D rendering tasks, so it was finally time to migrate PD's hoary old GDI+ interface to a more
' modern solution.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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

'When wrapping a DC, a surface needs to know the size of the object being painted on.  If an hWnd is supplied alongside
' the DC, we'll use that to auto-detect dimensions; otherwise, the caller needs to provide them.  (If the size is
' unknown, we'll use the size of the bitmap currently selected into the DC, but that's *not* reliable - so don't use it
' unless you know what you're doing!)
'
'This enum is only used internally.
Public Enum PD_2D_SIZE_DETECTION
    P2_SizeUnknown = 0
    P2_SizeFromHWnd = 1
    P2_SizeFromCaller = 2
End Enum

#If False Then
    Private Const P2_SizeUnknown = 0, P2_SizeFromHWnd = 1, P2_SizeFromCaller = 2
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

Public Enum PD_2D_CompositeMode
    P2_CM_Blend = 0
    P2_CM_Overwrite = 1
End Enum

#If False Then
    Private Const P2_CM_Blend = 0, P2_CM_Overwrite = 1
#End If

Public Enum PD_2D_DashCap
    P2_DC_Flat = 0
    P2_DC_Square = 1        'NOTE: GDI+ does not support square dash caps - only flat ones - so square simply remaps to flat
    P2_DC_Round = 2
    P2_DC_Triangle = 3
End Enum

#If False Then
    Private Const P2_DC_Flat = 0, P2_DC_Square = 1, P2_DC_Round = 2, P2_DC_Triangle = 3
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

Public Enum PD_2D_FileFormatImport
    P2_FFI_Undefined = -1
    P2_FFI_BMP = 0
    P2_FFI_ICO = 1
    P2_FFI_JPEG = 2
    P2_FFI_GIF = 3
    P2_FFI_PNG = 4
    P2_FFI_TIFF = 5
    P2_FFI_WMF = 6
    P2_FFI_EMF = 7
End Enum

#If False Then
    Private Const P2_FFI_Undefined = -1, P2_FFI_BMP = 0, P2_FFI_ICO = 1, P2_FFI_JPEG = 2, P2_FFI_GIF = 3, P2_FFI_PNG = 4, P2_FFI_TIFF = 5, P2_FFI_WMF = 6, P2_FFI_EMF = 7
#End If

Public Enum PD_2D_FileFormatExport
    P2_FFE_BMP = 0
    P2_FFE_GIF = 1
    P2_FFE_JPEG = 2
    P2_FFE_PNG = 3
    P2_FFE_TIFF = 4
End Enum

#If False Then
    Private Const P2_FFE_BMP = 0, P2_FFE_GIF = 1, P2_FFE_JPEG = 2, P2_FFE_PNG = 3, P2_FFE_TIFF = 4
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

Public Enum PD_2D_PixelOffset
    P2_PO_Normal = 0
    P2_PO_Half = 1
End Enum

#If False Then
    Private Const P2_PO_Normal = 0, P2_PO_Half = 1
#End If

Public Enum PD_2D_ResizeQuality
    P2_RQ_Fast = 0
    P2_RQ_Bilinear = 1
    P2_RQ_Bicubic = 2
End Enum

#If False Then
    Private Const P2_RQ_Fast = 0, P2_RQ_Bilinear = 1, P2_RQ_Bicubic = 2
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
    
    'IMPORTANT NOTE: clamp wrap mode does not work on all GDI+ calls; for example, it fails miserably
    ' on linear gradients for unknown reasons.  (See https://stackoverflow.com/questions/33225410/why-does-setting-lineargradientbrush-wrapmode-to-clamp-fail-with-argumentexcepti)
    P2_WM_Clamp = 4
End Enum

#If False Then
    Private Const P2_WM_Tile = 0, P2_WM_TileFlipX = 1, P2_WM_TileFlipY = 2, P2_WM_TileFlipXY = 3, P2_WM_Clamp = 4
#End If

'Certain structs are immensely helpful when drawing
Public Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Public Type PointFloat
    x As Single
    y As Single
End Type

Public Type PointLong
    x As Long
    y As Long
End Type

Public Type PointLong3D
    x As Long
    y As Long
    z As Long
End Type

Public Type RectL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RectL_WH
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Public Type RectF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Public Type SafeArrayBound
    cElements As Long
    lBound   As Long
End Type

Public Type SafeArray2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SafeArrayBound
End Type

Public Type SafeArray1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements As Long
    lBound   As Long
End Type

'PD's gradient format is straightforward, and it's declared here so functions can easily create their own gradient interfaces.
Public Type GradientPoint
    PointRGB As Long
    PointOpacity As Single
    PointPosition As Single
End Type

'If GDI+ is initialized successfully, this will be set to TRUE
Private m_GDIPlusAvailable As Boolean

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

Public Function GetNameOfFileFormat(ByVal srcFormat As PD_2D_FileFormatImport) As String
    Select Case srcFormat
        Case P2_FFI_BMP
            GetNameOfFileFormat = "BMP"
        Case P2_FFI_ICO
            GetNameOfFileFormat = "Icon"
        Case P2_FFI_JPEG
            GetNameOfFileFormat = "JPEG"
        Case P2_FFI_GIF
            GetNameOfFileFormat = "GIF"
        Case P2_FFI_PNG
            GetNameOfFileFormat = "PNG"
        Case P2_FFI_TIFF
            GetNameOfFileFormat = "TIFF"
        Case P2_FFI_WMF
            GetNameOfFileFormat = "WMF"
        Case P2_FFI_EMF
            GetNameOfFileFormat = "EMF"
        Case Else
            GetNameOfFileFormat = "Unknown file format"
    End Select
End Function

'Shortcut function for creating a new rectangular region with the default rendering backend
Public Function QuickCreateRegionRectangle(ByRef dstRegion As pd2DRegion, ByVal rLeft As Single, ByVal rTop As Single, ByVal rWidth As Single, ByVal rHeight As Single) As Boolean
    If (dstRegion Is Nothing) Then Set dstRegion = New pd2DRegion Else dstRegion.ResetAllProperties
    With dstRegion
        QuickCreateRegionRectangle = .AddRectangleF(rLeft, rTop, rWidth, rHeight, P2_CM_Replace)
    End With
End Function

'Shortcut function for quickly creating a blank surface with the default rendering backend and default rendering settings
Public Function QuickCreateBlankSurface(ByRef dstSurface As pd2DSurface, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, Optional ByVal surfaceSupportsAlpha As Boolean = True, Optional ByVal enableAntialiasing As Boolean = False, Optional ByVal initialColor As Long = vbWhite, Optional ByVal initialOpacity As Single = 100!) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateBlankSurface = .CreateBlankSurface(surfaceWidth, surfaceHeight, surfaceSupportsAlpha, initialColor, initialOpacity)
    End With
End Function

'Shortcut function for creating a new surface with the default rendering backend and default rendering settings
Public Function QuickCreateSurfaceFromDC(ByRef dstSurface As pd2DSurface, ByVal srcDC As Long, Optional ByVal enableAntialiasing As Boolean = False, Optional ByVal srcHWnd As Long = 0) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateSurfaceFromDC = .WrapSurfaceAroundDC(srcDC, srcHWnd)
    End With
End Function

Public Function QuickCreateSurfaceFromDIB(ByRef dstSurface As pd2DSurface, ByVal srcDIB As pdDIB, Optional ByVal enableAntialiasing As Boolean = False) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateSurfaceFromDIB = .WrapSurfaceAroundPDDIB(srcDIB)
    End With
End Function

Public Function QuickCreateSurfaceFromFile(ByRef dstSurface As pd2DSurface, ByVal srcPath As String) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        QuickCreateSurfaceFromFile = .CreateSurfaceFromFile(srcPath)
    End With
End Function

'Shortcut function for creating a solid brush
Public Function QuickCreateSolidBrush(ByRef dstBrush As pd2DBrush, Optional ByVal brushColor As Long = vbWhite, Optional ByVal brushOpacity As Single = 100!) As Boolean
    If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush Else dstBrush.ResetAllProperties
    With dstBrush
        .SetBrushColor brushColor
        .SetBrushOpacity brushOpacity
        QuickCreateSolidBrush = .CreateBrush()
    End With
End Function

'Shortcut function for creating a two-color gradient brush
Public Function QuickCreateTwoColorGradientBrush(ByRef dstBrush As pd2DBrush, ByRef gradientBoundary As RectF, Optional ByVal firstColor As Long = vbBlack, Optional ByVal secondColor As Long = vbWhite, Optional ByVal firstColorOpacity As Single = 100!, Optional ByVal secondColorOpacity As Single = 100!, Optional ByVal gradientShape As PD_2D_GradientShape = P2_GS_Linear, Optional ByVal gradientAngle As Single = 0!) As Boolean
    
    If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush Else dstBrush.ResetAllProperties
    
    Dim tmpGradient As pd2DGradient
    Set tmpGradient = New pd2DGradient
    With tmpGradient
        .SetGradientShape gradientShape
        .SetGradientAngle gradientAngle
        .CreateTwoPointGradient firstColor, secondColor, firstColorOpacity, secondColorOpacity
    End With
    
    With dstBrush
        .SetBrushMode P2_BM_Gradient
        .SetBoundaryRect gradientBoundary
        .SetBrushGradientAllSettings tmpGradient.GetGradientAsString
        QuickCreateTwoColorGradientBrush = .CreateBrush()
    End With
    
End Function

'Shortcut function for creating a solid pen
Public Function QuickCreateSolidPen(ByRef dstPen As pd2DPen, Optional ByVal penWidth As Single = 1!, Optional ByVal penColor As Long = vbWhite, Optional ByVal penOpacity As Single = 100!, Optional ByVal penLineJoin As PD_2D_LineJoin = P2_LJ_Miter, Optional ByVal penLineCap As PD_2D_LineCap = P2_LC_Flat) As Boolean
    If (dstPen Is Nothing) Then Set dstPen = New pd2DPen Else dstPen.ResetAllProperties
    With dstPen
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
        .SetPenWidth 3!
        .SetPenColor vbBlack
        .SetPenOpacity 75!
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreatePairOfUIPens = .CreatePen()
    End With
    
    With dstPenTop
        .SetPenWidth 1.6!
        If useHighlightColor Then .SetPenColor g_Themer.GetGenericUIColor(UI_Accent) Else .SetPenColor vbWhite
        .SetPenOpacity 87.5!
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreatePairOfUIPens = (QuickCreatePairOfUIPens And .CreatePen())
    End With
    
End Function

'LoadPicture replacement.  All pd2D interactions are handled internally, so you just pass the target object
' and source file path.
'
'The target object needs to have a DC property, or this function will fail.
Public Function QuickLoadPicture(ByRef dstObject As Object, ByVal srcPath As String, Optional ByVal resizeImageToFit As Boolean = True) As Boolean
    
    On Error GoTo LoadPictureFail
    
    Dim srcSurface As pd2DSurface
    If Drawing2D.QuickCreateSurfaceFromFile(srcSurface, srcPath) Then
        
        Dim dstSurface As pd2DSurface
        If Drawing2D.QuickCreateSurfaceFromDC(dstSurface, dstObject.hDC, True, dstObject.hWnd) Then
            
            If resizeImageToFit Then
                
                'If the source surface is smaller than the destination surface, center the image to fit
                If ((srcSurface.GetSurfaceWidth < dstSurface.GetSurfaceWidth) And (srcSurface.GetSurfaceHeight < dstSurface.GetSurfaceHeight)) Then
                    QuickLoadPicture = PD2D.DrawSurfaceI(dstSurface, (dstSurface.GetSurfaceWidth - srcSurface.GetSurfaceWidth) \ 2, (dstSurface.GetSurfaceHeight - srcSurface.GetSurfaceHeight) \ 2, srcSurface)
                Else
                
                    'Calculate the correct target size, and use that size when painting.
                    Dim newWidth As Long, newHeight As Long
                    PDMath.ConvertAspectRatio srcSurface.GetSurfaceWidth, srcSurface.GetSurfaceHeight, dstSurface.GetSurfaceWidth, dstSurface.GetSurfaceHeight, newWidth, newHeight
                    
                    dstSurface.SetSurfaceResizeQuality P2_RQ_Bicubic
                    QuickLoadPicture = PD2D.DrawSurfaceResizedI(dstSurface, (dstSurface.GetSurfaceWidth - newWidth) \ 2, (dstSurface.GetSurfaceHeight - newHeight) \ 2, newWidth, newHeight, srcSurface)
                    
                End If
                
            Else
                QuickLoadPicture = PD2D.DrawSurfaceI(dstSurface, 0, 0, srcSurface)
            End If
            
        End If
        
    End If
    
    Exit Function
    
LoadPictureFail:
    InternalError "QuickLoadPicture", Err.Description, Err.Number
    QuickLoadPicture = False
End Function

Public Function IsRenderingEngineActive(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            IsRenderingEngineActive = m_GDIPlusAvailable
        Case Else
            IsRenderingEngineActive = False
    End Select
End Function

'Start a new rendering backend
Public Function StartRenderingEngine(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean

    Select Case targetBackend
            
        Case P2_DefaultBackend, P2_GDIPlusBackend
            StartRenderingEngine = GDI_Plus.GDIP_StartEngine(False)
            m_GDIPlusAvailable = StartRenderingEngine
            
        Case Else
            InternalError "StartRenderingEngine", "unknown backend"
    
    End Select

End Function

'Stop a running rendering backend
Public Function StopRenderingEngine(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
        
    Select Case targetBackend
            
        Case P2_DefaultBackend, P2_GDIPlusBackend
            
            'Prior to release, see if any GDI+ object counts are non-zero; if they are, the caller needs to
            ' be notified, because those resources should be released before the backend disappears.
            If PD2D_DEBUG_MODE Then
                If (m_BrushCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_BrushCount_GDIPlus & " brush(es) active.  Release these objects before shutting down the drawing backend."
                If (m_PathCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_PathCount_GDIPlus & " path(s) active.  Release these objects before shutting down the drawing backend."
                If (m_PenCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_PenCount_GDIPlus & " pen(s) active.  Release these objects before shutting down the drawing backend."
                If (m_RegionCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_RegionCount_GDIPlus & " region(s) active.  Release these objects before shutting down the drawing backend."
                If (m_SurfaceCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_SurfaceCount_GDIPlus & " surface(s) active.  Release these objects before shutting down the drawing backend."
                If (m_TransformCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_TransformCount_GDIPlus & " transform(s) active.  Release these objects before shutting down the drawing backend."
            End If
            
            StopRenderingEngine = GDI_Plus.GDIP_StopEngine()
            m_GDIPlusAvailable = False
            
        Case Else
            InternalError "StopRenderingEngine", "unknown backend"
    
    End Select
    
End Function

'At present, Drawing2D errors are simply forwarded to the main error handler function at the bottom of this module.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    DEBUG_NotifyError "Drawing2D", errFunction, errDescription, errNum
End Sub

'DEBUG FUNCTIONS FOLLOW.  These functions should not be called directly.  They are invoked by other pd2D class when PD2D_DEBUG_MODE = TRUE.
Public Sub DEBUG_NotifyBrushCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_BrushCount_GDIPlus = m_BrushCount_GDIPlus + 1 Else m_BrushCount_GDIPlus = m_BrushCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyPathCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_PathCount_GDIPlus = m_PathCount_GDIPlus + 1 Else m_PathCount_GDIPlus = m_PathCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyPenCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_PenCount_GDIPlus = m_PenCount_GDIPlus + 1 Else m_PenCount_GDIPlus = m_PenCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyRegionCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_RegionCount_GDIPlus = m_RegionCount_GDIPlus + 1 Else m_RegionCount_GDIPlus = m_RegionCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifySurfaceCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus + 1 Else m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyTransformCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_TransformCount_GDIPlus = m_TransformCount_GDIPlus + 1 Else m_TransformCount_GDIPlus = m_TransformCount_GDIPlus - 1
End Sub

'In a default build, external pd2D classes relay any internal errors to this function.  You may wish to modify those classes
' to raise their own error events, or perhaps handle their errors internally.  (By default, pd2D does *not* halt on errors.)
Public Sub DEBUG_NotifyError(ByRef errClassName As String, ByRef errFunctionName As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    PDDebug.LogAction "WARNING!  pd2D error in " & errClassName & "." & errFunctionName & ": " & errDescription
    If (errNum <> 0) Then PDDebug.LogAction "  (If it helps, an error number was also reported: #" & errNum & ")"
End Sub

'These functions exist to help with XML serialization.  They use consistent names against
' the current enums, and if the enums ever change order in the future, existing XML strings
' will still produce correct results.
Public Function XML_GetNameOfWrapMode(ByVal srcWrapMode As PD_2D_WrapMode) As String
    Select Case srcWrapMode
        Case P2_WM_Tile
            XML_GetNameOfWrapMode = "tile"
        Case P2_WM_TileFlipX
            XML_GetNameOfWrapMode = "tile-flip-x"
        Case P2_WM_TileFlipY
            XML_GetNameOfWrapMode = "tile-flip-y"
        Case P2_WM_TileFlipXY
            XML_GetNameOfWrapMode = "tile-flip-xy"
        Case P2_WM_Clamp
            XML_GetNameOfWrapMode = "clamp"
        Case Else
            XML_GetNameOfWrapMode = "tile"
    End Select
End Function

Public Function XML_GetWrapModeFromName(ByRef srcName As String) As PD_2D_WrapMode
    Select Case srcName
        Case "tile"
            XML_GetWrapModeFromName = P2_WM_Tile
        Case "tile-flip-x"
            XML_GetWrapModeFromName = P2_WM_TileFlipX
        Case "tile-flip-y"
            XML_GetWrapModeFromName = P2_WM_TileFlipY
        Case "tile-flip-xy"
            XML_GetWrapModeFromName = P2_WM_TileFlipXY
        Case "clamp"
            XML_GetWrapModeFromName = P2_WM_Clamp
        Case Else
            XML_GetWrapModeFromName = P2_WM_Tile
    End Select
End Function

Public Function XML_GetNameOfBrushMode(ByVal srcBrushMode As PD_2D_BrushMode) As String
    Select Case srcBrushMode
        Case P2_BM_Solid
            XML_GetNameOfBrushMode = "solid"
        Case P2_BM_Pattern
            XML_GetNameOfBrushMode = "pattern"
        Case P2_BM_Gradient
            XML_GetNameOfBrushMode = "gradient"
        Case P2_BM_Texture
            XML_GetNameOfBrushMode = "texture"
        Case Else
            XML_GetNameOfBrushMode = "solid"
    End Select
End Function

Public Function XML_GetBrushModeFromName(ByRef srcName As String) As PD_2D_BrushMode
    Select Case srcName
        Case "solid"
            XML_GetBrushModeFromName = P2_BM_Solid
        Case "pattern"
            XML_GetBrushModeFromName = P2_BM_Pattern
        Case "gradient"
            XML_GetBrushModeFromName = P2_BM_Gradient
        Case "texture"
            XML_GetBrushModeFromName = P2_BM_Texture
        Case Else
            XML_GetBrushModeFromName = P2_BM_Solid
    End Select
End Function

Public Function XML_GetNameOfPattern(ByVal srcPattern As PD_2D_PatternStyle) As String
    Select Case srcPattern
        Case P2_PS_Horizontal
            XML_GetNameOfPattern = "x"
        Case P2_PS_Vertical
            XML_GetNameOfPattern = "y"
        Case P2_PS_ForwardDiagonal
            XML_GetNameOfPattern = "forward-dg"
        Case P2_PS_BackwardDiagonal
            XML_GetNameOfPattern = "backward-dg"
        Case P2_PS_Cross
            XML_GetNameOfPattern = "cross"
        Case P2_PS_DiagonalCross
            XML_GetNameOfPattern = "dg-cross"
        Case P2_PS_05Percent
            XML_GetNameOfPattern = "pc-05"
        Case P2_PS_10Percent
            XML_GetNameOfPattern = "pc-10"
        Case P2_PS_20Percent
            XML_GetNameOfPattern = "pc-20"
        Case P2_PS_25Percent
            XML_GetNameOfPattern = "pc-25"
        Case P2_PS_30Percent
            XML_GetNameOfPattern = "pc-30"
        Case P2_PS_40Percent
            XML_GetNameOfPattern = "pc-40"
        Case P2_PS_50Percent
            XML_GetNameOfPattern = "pc-50"
        Case P2_PS_60Percent
            XML_GetNameOfPattern = "pc-60"
        Case P2_PS_70Percent
            XML_GetNameOfPattern = "pc-70"
        Case P2_PS_75Percent
            XML_GetNameOfPattern = "pc-75"
        Case P2_PS_80Percent
            XML_GetNameOfPattern = "pc-80"
        Case P2_PS_90Percent
            XML_GetNameOfPattern = "pc-90"
        Case P2_PS_LightDownwardDiagonal
            XML_GetNameOfPattern = "light-down-dg"
        Case P2_PS_LightUpwardDiagonal
            XML_GetNameOfPattern = "light-up-dg"
        Case P2_PS_DarkDownwardDiagonal
            XML_GetNameOfPattern = "dark-down-dg"
        Case P2_PS_DarkUpwardDiagonal
            XML_GetNameOfPattern = "dark-up-dg"
        Case P2_PS_WideDownwardDiagonal
            XML_GetNameOfPattern = "wide-down-dg"
        Case P2_PS_WideUpwardDiagonal
            XML_GetNameOfPattern = "wide-up-dg"
        Case P2_PS_LightVertical
            XML_GetNameOfPattern = "light-y"
        Case P2_PS_LightHorizontal
            XML_GetNameOfPattern = "light-x"
        Case P2_PS_NarrowVertical
            XML_GetNameOfPattern = "narrow-y"
        Case P2_PS_NarrowHorizontal
            XML_GetNameOfPattern = "narrow-x"
        Case P2_PS_DarkVertical
            XML_GetNameOfPattern = "dark-y"
        Case P2_PS_DarkHorizontal
            XML_GetNameOfPattern = "dark-x"
        Case P2_PS_DashedDownwardDiagonal
            XML_GetNameOfPattern = "dash-down-dg"
        Case P2_PS_DashedUpwardDiagonal
            XML_GetNameOfPattern = "dash-up-dg"
        Case P2_PS_DashedHorizontal
            XML_GetNameOfPattern = "dash-x"
        Case P2_PS_DashedVertical
            XML_GetNameOfPattern = "dash-y"
        Case P2_PS_SmallConfetti
            XML_GetNameOfPattern = "confetti-s"
        Case P2_PS_LargeConfetti
            XML_GetNameOfPattern = "confetti-l"
        Case P2_PS_ZigZag
            XML_GetNameOfPattern = "zigzag"
        Case P2_PS_Wave
            XML_GetNameOfPattern = "wave"
        Case P2_PS_DiagonalBrick
            XML_GetNameOfPattern = "brick-dg"
        Case P2_PS_HorizontalBrick
            XML_GetNameOfPattern = "brick-x"
        Case P2_PS_Weave
            XML_GetNameOfPattern = "weave"
        Case P2_PS_Plaid
            XML_GetNameOfPattern = "plaid"
        Case P2_PS_Divot
            XML_GetNameOfPattern = "divot"
        Case P2_PS_DottedGrid
            XML_GetNameOfPattern = "dot-grid"
        Case P2_PS_DottedDiamond
            XML_GetNameOfPattern = "dot-diamond"
        Case P2_PS_Shingle
            XML_GetNameOfPattern = "shingle"
        Case P2_PS_Trellis
            XML_GetNameOfPattern = "trellis"
        Case P2_PS_Sphere
            XML_GetNameOfPattern = "sphere"
        Case P2_PS_SmallGrid
            XML_GetNameOfPattern = "grid-s"
        Case P2_PS_SmallCheckerBoard
            XML_GetNameOfPattern = "checker-s"
        Case P2_PS_LargeCheckerBoard
            XML_GetNameOfPattern = "checker-l"
        Case P2_PS_OutlinedDiamond
            XML_GetNameOfPattern = "diamond-outline"
        Case P2_PS_SolidDiamond
            XML_GetNameOfPattern = "diamond-solid"
        Case Else
            XML_GetNameOfPattern = "x"
    End Select
End Function

Public Function XML_GetPatternFromName(ByRef srcName As String) As PD_2D_PatternStyle
    Select Case srcName
        Case "x"
            XML_GetPatternFromName = P2_PS_Horizontal
        Case "y"
            XML_GetPatternFromName = P2_PS_Vertical
        Case "forward-dg"
            XML_GetPatternFromName = P2_PS_ForwardDiagonal
        Case "backward-dg"
            XML_GetPatternFromName = P2_PS_BackwardDiagonal
        Case "cross"
            XML_GetPatternFromName = P2_PS_Cross
        Case "dg-cross"
            XML_GetPatternFromName = P2_PS_DiagonalCross
        Case "pc-05"
            XML_GetPatternFromName = P2_PS_05Percent
        Case "pc-10"
            XML_GetPatternFromName = P2_PS_10Percent
        Case "pc-20"
            XML_GetPatternFromName = P2_PS_20Percent
        Case "pc-25"
            XML_GetPatternFromName = P2_PS_25Percent
        Case "pc-30"
            XML_GetPatternFromName = P2_PS_30Percent
        Case "pc-40"
            XML_GetPatternFromName = P2_PS_40Percent
        Case "pc-50"
            XML_GetPatternFromName = P2_PS_50Percent
        Case "pc-60"
            XML_GetPatternFromName = P2_PS_60Percent
        Case "pc-70"
            XML_GetPatternFromName = P2_PS_70Percent
        Case "pc-75"
            XML_GetPatternFromName = P2_PS_75Percent
        Case "pc-80"
            XML_GetPatternFromName = P2_PS_80Percent
        Case "pc-90"
            XML_GetPatternFromName = P2_PS_90Percent
        Case "light-down-dg"
            XML_GetPatternFromName = P2_PS_LightDownwardDiagonal
        Case "light-up-dg"
            XML_GetPatternFromName = P2_PS_LightUpwardDiagonal
        Case "dark-down-dg"
            XML_GetPatternFromName = P2_PS_DarkDownwardDiagonal
        Case "dark-up-dg"
            XML_GetPatternFromName = P2_PS_DarkUpwardDiagonal
        Case "wide-down-dg"
            XML_GetPatternFromName = P2_PS_WideDownwardDiagonal
        Case "wide-up-dg"
            XML_GetPatternFromName = P2_PS_WideUpwardDiagonal
        Case "light-y"
            XML_GetPatternFromName = P2_PS_LightVertical
        Case "light-x"
            XML_GetPatternFromName = P2_PS_LightHorizontal
        Case "narrow-y"
            XML_GetPatternFromName = P2_PS_NarrowVertical
        Case "narrow-x"
            XML_GetPatternFromName = P2_PS_NarrowHorizontal
        Case "dark-y"
            XML_GetPatternFromName = P2_PS_DarkVertical
        Case "dark-x"
            XML_GetPatternFromName = P2_PS_DarkHorizontal
        Case "dash-down-dg"
            XML_GetPatternFromName = P2_PS_DashedDownwardDiagonal
        Case "dash-up-dg"
            XML_GetPatternFromName = P2_PS_DashedUpwardDiagonal
        Case "dash-x"
            XML_GetPatternFromName = P2_PS_DashedHorizontal
        Case "dash-y"
            XML_GetPatternFromName = P2_PS_DashedVertical
        Case "confetti-s"
            XML_GetPatternFromName = P2_PS_SmallConfetti
        Case "confetti-l"
            XML_GetPatternFromName = P2_PS_LargeConfetti
        Case "zigzag"
            XML_GetPatternFromName = P2_PS_ZigZag
        Case "wave"
            XML_GetPatternFromName = P2_PS_Wave
        Case "brick-dg"
            XML_GetPatternFromName = P2_PS_DiagonalBrick
        Case "brick-x"
            XML_GetPatternFromName = P2_PS_HorizontalBrick
        Case "weave"
            XML_GetPatternFromName = P2_PS_Weave
        Case "plaid"
            XML_GetPatternFromName = P2_PS_Plaid
        Case "divot"
            XML_GetPatternFromName = P2_PS_Divot
        Case "dot-grid"
            XML_GetPatternFromName = P2_PS_DottedGrid
        Case "dot-diamond"
            XML_GetPatternFromName = P2_PS_DottedDiamond
        Case "shingle"
            XML_GetPatternFromName = P2_PS_Shingle
        Case "trellis"
            XML_GetPatternFromName = P2_PS_Trellis
        Case "sphere"
            XML_GetPatternFromName = P2_PS_Sphere
        Case "grid-s"
            XML_GetPatternFromName = P2_PS_SmallGrid
        Case "checker-s"
            XML_GetPatternFromName = P2_PS_SmallCheckerBoard
        Case "checker-l"
            XML_GetPatternFromName = P2_PS_LargeCheckerBoard
        Case "diamond-outline"
            XML_GetPatternFromName = P2_PS_OutlinedDiamond
        Case "diamond-solid"
            XML_GetPatternFromName = P2_PS_SolidDiamond
        Case Else
            XML_GetPatternFromName = P2_PS_Horizontal
    End Select
End Function

Public Function XML_GetNameOfGradientShape(ByVal srcShape As PD_2D_GradientShape) As String
    Select Case srcShape
        Case P2_GS_Linear
            XML_GetNameOfGradientShape = "linear"
        Case P2_GS_Reflection
            XML_GetNameOfGradientShape = "reflect"
        Case P2_GS_Radial
            XML_GetNameOfGradientShape = "radial"
        Case P2_GS_Rectangle
            XML_GetNameOfGradientShape = "rectangle"
        Case P2_GS_Diamond
            XML_GetNameOfGradientShape = "diamond"
        Case Else
            XML_GetNameOfGradientShape = "linear"
    End Select
End Function

Public Function XML_GetGradientShapeFromName(ByRef srcName As String) As PD_2D_GradientShape
    Select Case srcName
        Case "linear"
            XML_GetGradientShapeFromName = P2_GS_Linear
        Case "reflect"
            XML_GetGradientShapeFromName = P2_GS_Reflection
        Case "radial"
            XML_GetGradientShapeFromName = P2_GS_Radial
        Case "rectangle"
            XML_GetGradientShapeFromName = P2_GS_Rectangle
        Case "diamond"
            XML_GetGradientShapeFromName = P2_GS_Diamond
        Case Else
            XML_GetGradientShapeFromName = P2_GS_Linear
    End Select
End Function

Public Function XML_GetNameOfLineCap(ByVal srcLineCap As PD_2D_LineCap) As String
    Select Case srcLineCap
        Case P2_LC_Flat
            XML_GetNameOfLineCap = "flat"
        Case P2_LC_Square
            XML_GetNameOfLineCap = "square"
        Case P2_LC_Round
            XML_GetNameOfLineCap = "round"
        Case P2_LC_Triangle
            XML_GetNameOfLineCap = "triangle"
        Case P2_LC_FlatAnchor
            XML_GetNameOfLineCap = "anchor-flat"
        Case P2_LC_SquareAnchor
            XML_GetNameOfLineCap = "anchor-square"
        Case P2_LC_RoundAnchor
            XML_GetNameOfLineCap = "anchor-round"
        Case P2_LC_DiamondAnchor
            XML_GetNameOfLineCap = "anchor-diamond"
        Case P2_LC_ArrowAnchor
            XML_GetNameOfLineCap = "anchor-arrow"
        Case P2_LC_Custom
            XML_GetNameOfLineCap = "custom"
        Case Else
            XML_GetNameOfLineCap = "flat"
    End Select
End Function

Public Function XML_GetLineCapFromName(ByRef srcName As String) As PD_2D_LineCap
    Select Case srcName
        Case "flat"
            XML_GetLineCapFromName = P2_LC_Flat
        Case "square"
            XML_GetLineCapFromName = P2_LC_Square
        Case "round"
            XML_GetLineCapFromName = P2_LC_Round
        Case "triangle"
            XML_GetLineCapFromName = P2_LC_Triangle
        Case "anchor-flat"
            XML_GetLineCapFromName = P2_LC_FlatAnchor
        Case "anchor-square"
            XML_GetLineCapFromName = P2_LC_SquareAnchor
        Case "anchor-round"
            XML_GetLineCapFromName = P2_LC_RoundAnchor
        Case "anchor-diamond"
            XML_GetLineCapFromName = P2_LC_DiamondAnchor
        Case "anchor-arrow"
            XML_GetLineCapFromName = P2_LC_ArrowAnchor
        Case "custom"
            XML_GetLineCapFromName = P2_LC_Custom
        Case Else
            XML_GetLineCapFromName = P2_LC_Flat
    End Select
End Function

Public Function XML_GetNameOfDashCap(ByVal srcDashCap As PD_2D_DashCap) As String
    Select Case srcDashCap
        Case P2_DC_Flat
            XML_GetNameOfDashCap = "flat"
        Case P2_DC_Square
            XML_GetNameOfDashCap = "square"
        Case P2_DC_Round
            XML_GetNameOfDashCap = "round"
        Case P2_DC_Triangle
            XML_GetNameOfDashCap = "triangle"
        Case Else
            XML_GetNameOfDashCap = "flat"
    End Select
End Function

Public Function XML_GetDashCapFromName(ByRef srcName As String) As PD_2D_DashCap
    Select Case srcName
        Case "flat"
            XML_GetDashCapFromName = P2_DC_Flat
        Case "square"
            XML_GetDashCapFromName = P2_DC_Square
        Case "round"
            XML_GetDashCapFromName = P2_DC_Round
        Case "triangle"
            XML_GetDashCapFromName = P2_DC_Triangle
        Case Else
            XML_GetDashCapFromName = P2_DC_Flat
    End Select
End Function

Public Function XML_GetNameOfLineJoin(ByVal srcLineJoin As PD_2D_LineJoin) As String
    Select Case srcLineJoin
        Case P2_LJ_Miter
            XML_GetNameOfLineJoin = "miter"
        Case P2_LJ_Bevel
            XML_GetNameOfLineJoin = "bevel"
        Case P2_LJ_Round
            XML_GetNameOfLineJoin = "round"
        Case Else
            XML_GetNameOfLineJoin = "miter"
    End Select
End Function

Public Function XML_GetLineJoinFromName(ByRef srcName As String) As PD_2D_LineJoin
    Select Case srcName
        Case "miter"
            XML_GetLineJoinFromName = P2_LJ_Miter
        Case "bevel"
            XML_GetLineJoinFromName = P2_LJ_Bevel
        Case "round"
            XML_GetLineJoinFromName = P2_LJ_Round
        Case Else
            XML_GetLineJoinFromName = P2_LJ_Miter
    End Select
End Function

Public Function XML_GetNameOfDashStyle(ByVal srcPenStyle As PD_2D_DashStyle) As String
    Select Case srcPenStyle
        Case P2_DS_Solid
            XML_GetNameOfDashStyle = "solid"
        Case P2_DS_Dash
            XML_GetNameOfDashStyle = "dash"
        Case P2_DS_Dot
            XML_GetNameOfDashStyle = "dot"
        Case P2_DS_DashDot
            XML_GetNameOfDashStyle = "dash-dot"
        Case P2_DS_DashDotDot
            XML_GetNameOfDashStyle = "dash-dot-dot"
        Case P2_DS_Custom
            XML_GetNameOfDashStyle = "custom"
        Case Else
            XML_GetNameOfDashStyle = "solid"
    End Select
End Function

Public Function XML_GetDashStyleFromName(ByRef srcName As String) As PD_2D_DashStyle
    Select Case srcName
        Case "solid"
            XML_GetDashStyleFromName = P2_DS_Solid
        Case "dash"
            XML_GetDashStyleFromName = P2_DS_Dash
        Case "dot"
            XML_GetDashStyleFromName = P2_DS_Dot
        Case "dash-dot"
            XML_GetDashStyleFromName = P2_DS_DashDot
        Case "dash-dot-dot"
            XML_GetDashStyleFromName = P2_DS_DashDotDot
        Case "custom"
            XML_GetDashStyleFromName = P2_DS_Custom
        Case Else
            XML_GetDashStyleFromName = P2_DS_Solid
    End Select
End Function

