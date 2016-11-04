Attribute VB_Name = "Paintbrush"
'***************************************************************************
'Paintbrush tool interface
'Copyright 2016-2016 by Tanner Helland
'Created: 1/November/16
'Last updated: 1/November/16
'Last update: initial build
'
'To simplify the design of the primary canvas, it makes brush-related requests to this module.  This module
' then handles all the messy business of managing the actual background brush data.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Internally, we switch between different brush rendering engines depending on the current brush settings.
' The caller doesn't need to concern themselves with this; it's used only to determine internal rendering paths.
Private Enum BRUSH_ENGINE
    BE_GDIPlus = 0
    BE_PhotoDemon = 1
End Enum

#If False Then
    Private Const BE_GDIPlus = 0, BE_PhotoDemon = 1
#End If

Public Enum BRUSH_SOURCES
    BS_Color = 0
End Enum

#If False Then
    Private Const BS_Color = 0
#End If

Public Enum BRUSH_ATTRIBUTES
    BA_Source = 0
    BA_Radius = 1
    BA_Opacity = 2
    BA_Blendmode = 3
    
    'Source-specific values can be stored here, as necessary
    BA_SourceColor = 1000
End Enum

#If False Then
    Private Const BA_Source = 0, BA_Radius = 1, BA_Opacity = 2, BA_Blendmode = 3, BA_SourceColor = 1000
#End If

'The current brush engine is stored here.  Note that this value is not correct until a call has been made to
' the CreateCurrentBrush() function; this function searches brush attributes and determines which brush engine
' to use.
Private m_BrushEngine As BRUSH_ENGINE
Private m_BrushOutlineImage As pdDIB, m_BrushOutlinePath As pd2DPath

'Brush resources, used only as necessary.  Check for null values before using.
Private m_GDIPPen As pd2DPen

'Brush attributes are stored in these variables
Private m_BrushSource As BRUSH_SOURCES
Private m_BrushRadius As Single
Private m_BrushOpacity As Single
Private m_BrushBlendmode As LAYER_BLENDMODE

'Note that some brush attributes only exist for certain brush sources.
Private m_BrushSourceColor As Long

'If brush properties have changed since the last brush creation, this is set to FALSE.  We use this to optimize
' brush creation behavior.
Private m_BrushIsReady As Boolean
Private m_BrushCreatedAtLeastOnce As Boolean

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform any
' special tracking calculations.
Private m_MouseDown As Boolean
Private m_MouseX As Single, m_MouseY As Single

Public Function GetBrushSource() As BRUSH_SOURCES
    GetBrushSource = m_BrushSource
End Function

Public Function GetBrushRadius() As Single
    GetBrushRadius = m_BrushRadius
End Function

Public Function GetBrushOpacity() As Single
    GetBrushOpacity = m_BrushOpacity
End Function

Public Function GetBrushBlendMode() As LAYER_BLENDMODE
    GetBrushBlendMode = m_BrushBlendmode
End Function

Public Function GetBrushSourceColor() As Long
    GetBrushSourceColor = m_BrushSourceColor
End Function

Public Sub SetBrushSource(ByVal newSource As BRUSH_SOURCES)
    If (newSource <> m_BrushSource) Then
        m_BrushSource = newSource
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushRadius(ByVal newRadius As Single)
    If (newRadius <> m_BrushRadius) Then
        m_BrushRadius = newRadius
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushOpacity(Optional ByVal newOpacity As Single = 100#)
    If (newOpacity <> m_BrushOpacity) Then
        m_BrushOpacity = newOpacity
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushBlendMode(Optional ByVal newBlendMode As LAYER_BLENDMODE = BL_NORMAL)
    If (newBlendMode <> m_BrushBlendmode) Then
        m_BrushBlendmode = newBlendMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSourceColor(Optional ByVal newColor As Long = vbWhite)
    If (newColor <> m_BrushSourceColor) Then
        m_BrushSourceColor = newColor
        m_BrushIsReady = False
    End If
End Sub

Public Function GetBrushProperty(ByVal bProperty As BRUSH_ATTRIBUTES) As Variant
    
    Select Case bProperty
        Case BA_Source
            GetBrushProperty = GetBrushSource()
        Case BA_Radius
            GetBrushProperty = GetBrushRadius()
        Case BA_Opacity
            GetBrushProperty = GetBrushOpacity()
        Case BA_Blendmode
            GetBrushProperty = GetBrushBlendMode()
        Case BA_SourceColor
            GetBrushProperty = GetBrushSourceColor()
    End Select
    
End Function

Public Sub SetBrushProperty(ByVal bProperty As BRUSH_ATTRIBUTES, ByVal newPropValue As Variant)
    
    Select Case bProperty
        Case BA_Source
            SetBrushSource newPropValue
        Case BA_Radius
            SetBrushRadius newPropValue
        Case BA_Opacity
            SetBrushOpacity newPropValue
        Case BA_Blendmode
            SetBrushBlendMode newPropValue
        Case BA_SourceColor
            SetBrushSourceColor newPropValue
    End Select
    
End Sub

Public Sub CreateCurrentBrush(Optional ByVal alsoCreateBrushOutline As Boolean = True, Optional ByVal forceCreation As Boolean = False)
    
    If ((Not m_BrushIsReady) Or forceCreation Or (Not m_BrushCreatedAtLeastOnce)) Then
    
        'In the future we'll be implementing a full custom brush engine, but for this early testing phase,
        ' I'm restricting things to GDI+ for simplicity's sake.
        m_BrushEngine = BE_GDIPlus
        
        Select Case m_BrushEngine
            
            Case BE_GDIPlus
                'For now, create a circular pen at the current radius
                If (m_GDIPPen Is Nothing) Then Set m_GDIPPen = New pd2DPen
                Drawing2D.QuickCreateSolidPen m_GDIPPen
        
        End Select
        
        'Whenever we create a new brush, we should also refresh the current brush outline
        If alsoCreateBrushOutline Then CreateCurrentBrushOutline
        
        m_BrushIsReady = True
        m_BrushCreatedAtLeastOnce = True
        
    End If
    
End Sub

'As part of rendering the current brush, we also need to render a brush outline onto the canvas at the current
' mouse location.  The specific outline technique used varies by brush engine.
Public Sub CreateCurrentBrushOutline()
    
    Select Case m_BrushEngine
    
        'If this is a GDI+ brush, outline creation is pretty easy.  Assume a circular brush and simply
        ' create a path at that same radius.
        Case BE_GDIPlus
        
            Set m_BrushOutlinePath = New pd2DPath
            If (m_BrushRadius > 0#) Then m_BrushOutlinePath.AddCircle 0, 0, m_BrushRadius / 2 + 1#
    
    End Select

End Sub

'Notify the brush engine of the current mouse position.  Coordinates should always be in *image* coordinate space,
' not screen space.  (Translation between spaces will be handled internally.)
Public Sub NotifyBrushXY(ByVal mouseButtonDown As Boolean, ByVal srcX As Single, ByVal srcY As Single)
    
    Dim isFirstClick As Boolean
    
    'If this is a MouseDown operation, we need to prep the full paint engine.
    ' (TODO: initialize this elsewhere, so there's no "hesitation" on first paint.)
    If ((Not m_MouseDown) And mouseButtonDown) Then
        
        isFirstClick = True
        
        'Make sure the current scratch layer is properly initialized
        pdImages(g_CurrentImage).ResetScratchLayer
        
        'Reset the "last mouse position" values to match the current ones
        m_MouseX = srcX
        m_MouseY = srcY
    
    End If
    
    'If the mouse button is down, perform painting between the old and new points.
    ' (All painting occurs in image coordinate space, and is applied to the current image's scratch layer.)
    If mouseButtonDown Then
    
        'Create required pd2D drawing tools (a painter and surface)
        Dim cPainter As pd2DPainter
        Drawing2D.QuickCreatePainter cPainter
        
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBDC, True
        
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, m_BrushRadius, m_BrushSourceColor, , P2_LJ_Round, P2_LC_Round
        
        'Render the line
        'If isFirstClick Then
        '    cPainter.DrawCircleF cSurface, cPen, srcX, srcY, m_BrushRadius \ 2
        'Else
            cPainter.DrawLineF cSurface, cPen, m_MouseX, m_MouseY, srcX, srcY
        'End If
        
        Set cPainter = Nothing: Set cSurface = Nothing: Set cPen = Nothing
    
        pdImages(g_CurrentImage).ScratchLayer.NotifyOfDestructiveChanges
    
    End If
    
    'With all painting tasks complete, update all old state values to match the new state values
    m_MouseDown = mouseButtonDown
    m_MouseX = srcX
    m_MouseY = srcY
    
End Sub

'Render the current brush outline to the canvas, using the stored mouse coordinates as the brush's position
Public Sub RenderBrushOutline(ByRef targetCanvas As pdCanvas)
    
    'If a brush outline doesn't exist, create one now
    If (Not m_BrushIsReady) Then CreateCurrentBrush True
    
    'Start by creating a transformation from the image space to the canvas space
    Dim canvasMatrix As pd2DTransform
    Drawing.GetTransformFromImageToCanvas canvasMatrix, targetCanvas, pdImages(g_CurrentImage), m_MouseX, m_MouseY
    
    'We also want to pinpoint the precise cursor position
    Dim cursX As Double, cursY As Double
    Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, pdImages(g_CurrentImage), m_MouseX, m_MouseY, cursX, cursY
    
    'If the on-screen brush size is above a certain threshold, we'll paint a full brush outline.
    ' If it's too small, we'll only paint a cross in the current brush position.
    Dim onScreenSize As Double
    onScreenSize = Drawing.ConvertImageSizeToCanvasSize(m_BrushRadius, pdImages(g_CurrentImage))
    
    Dim brushTooSmall As Boolean
    If (onScreenSize < 5) Then brushTooSmall = True
    
    'Create a pair of UI pens
    Dim innerPen As pd2DPen, outerPen As pd2DPen
    Drawing2D.QuickCreatePairOfUIPens outerPen, innerPen
    
    'Create other required pd2D drawing tools (a painter and surface)
    Dim cPainter As pd2DPainter
    Drawing2D.QuickCreatePainter cPainter
    
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetCanvas.hDC, True
    
    'Regardless of brush size, paint a target cursor
    Dim crossLength As Single, outerCrossBorder As Single
    crossLength = 3#
    outerCrossBorder = 0.5
    
    cPainter.DrawLineF cSurface, outerPen, cursX, cursY - crossLength - outerCrossBorder, cursX, cursY + crossLength + outerCrossBorder
    cPainter.DrawLineF cSurface, outerPen, cursX - crossLength - outerCrossBorder, cursY, cursX + crossLength + outerCrossBorder, cursY
    cPainter.DrawLineF cSurface, innerPen, cursX, cursY - crossLength, cursX, cursY + crossLength
    cPainter.DrawLineF cSurface, innerPen, cursX - crossLength, cursY, cursX + crossLength, cursY
    
    'If size allows, render a transformed brush outline onto the canvas as well
    If (Not brushTooSmall) Then
        
        'Get a copy of the current brush outline, transformed into position
        Dim copyOfBrushOutline As pd2DPath
        Set copyOfBrushOutline = New pd2DPath
        copyOfBrushOutline.CloneExistingPath m_BrushOutlinePath
        copyOfBrushOutline.ApplyTransformation canvasMatrix
    
        cPainter.DrawPath cSurface, outerPen, copyOfBrushOutline
        cPainter.DrawPath cSurface, innerPen, copyOfBrushOutline
    End If
    
    Set cPainter = Nothing: Set cSurface = Nothing
    Set innerPen = Nothing: Set outerPen = Nothing
    
End Sub

'A brush is considered active if the mouse state is currently DOWN, or if it is up but we are still rendering a
' previous stroke.
Public Function IsBrushActive() As Boolean
    IsBrushActive = m_MouseDown
End Function

'Before PD closes, you *must* call this function!  It will free any lingering brush resources (which are cached
' for performance reasons).
Public Sub FreeBrushResources()

    Set m_GDIPPen = Nothing

End Sub
