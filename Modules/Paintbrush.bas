Attribute VB_Name = "Paintbrush"
'***************************************************************************
'Paintbrush tool interface
'Copyright 2016-2018 by Tanner Helland
'Created: 1/November/16
'Last updated: 15/December/16
'Last update: ongoing performance improvements
'
'To simplify the design of the primary canvas, it makes brush-related requests to this module.  This module
' then handles all the messy business of managing the actual background brush data.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'PD allows the paintbrush to explicity set its own quality settings, independent of what the viewport quality
' settings are.  This was very helpful during debugging (to isolate performance bottlenecks), but now that
' things are working well, I prefer to lock the paintbrush renderer to the same settings as the viewport.
' If you set this to TRUE, make sure the m_BrushPreviewQuality variable is initialized accordingly.
Private Const USE_PAINTBRUSH_DEBUG_QUALITIES As Boolean = False

'Internally, we switch between different brush rendering engines depending on the current brush settings.
' The caller doesn't need to concern themselves with this; it's used only to determine internal rendering paths.
Private Enum PD_BrushEngine
    BE_GDIPlus = 0
    BE_PhotoDemon = 1
End Enum

#If False Then
    Private Const BE_GDIPlus = 0, BE_PhotoDemon = 1
#End If

Public Enum PD_BrushSource
    BS_Color = 0
End Enum

#If False Then
    Private Const BS_Color = 0
#End If

Public Enum PD_BrushStyle
    BS_Pencil = 0
    BS_SoftBrush = 1
End Enum

#If False Then
    Private Const BS_Pencil = 0, BS_SoftBrush = 1
#End If

Public Enum PD_BrushAttributes
    BA_Source = 0
    BA_Style = 1
    BA_Size = 2
    BA_Opacity = 3
    BA_BlendMode = 4
    BA_AlphaMode = 5
    BA_Antialiasing = 6
    BA_Hardness = 7
    BA_Spacing = 8
    BA_Flow = 9
    
    'Source-specific values can be stored here, as relevant
    BA_SourceColor = 1000
End Enum

#If False Then
    Private Const BA_Source = 0, BA_Style = 1, BA_Size = 2, BA_Opacity = 3, BA_BlendMode = 4, BA_AlphaMode = 5, BA_Antialiasing = 6
    Private Const BA_Hardness = 7, BA_Spacing = 8, BA_Flow = 9
    Private Const BA_SourceColor = 1000
#End If

'The current brush engine is stored here.  Note that this value is not correct until a call has been made to
' the CreateCurrentBrush() function; this function searches brush attributes and determines which brush engine
' to use.
Private m_BrushEngine As PD_BrushEngine
Private m_BrushOutlineImage As pdDIB, m_BrushOutlinePath As pd2DPath

'Brush preview quality.  At present, this is directly exposed on the paintbrush toolpanel.  This may change
' in the future, but for now, it's very helpful for testing.
Private m_BrushPreviewQuality As PD_PerformanceSetting

'Brush resources, used only as necessary.  Check for null values before using.
Private m_GDIPPen As pd2DPen
Private m_CustomPenImage As pd2DSurface, m_SrcPenDIB As pdDIB

'Brush attributes are stored in these variables
Private m_BrushSource As PD_BrushSource
Private m_BrushStyle As PD_BrushStyle
Private m_BrushSize As Single
Private m_BrushOpacity As Single
Private m_BrushBlendmode As PD_BlendMode
Private m_BrushAlphamode As PD_AlphaMode
Private m_BrushAntialiasing As PD_2D_Antialiasing
Private m_BrushHardness As Single
Private m_BrushSpacing As Single
Private m_BrushFlow As Single

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

'Brush dynamics are calculated on-the-fly, and they include things like velocity, distance, angle, and more.
Private m_DistPixels As Long, m_BrushSizeInt As Long
Private m_BrushSpacingCheck As Long

'As brush movements are relayed to us, we keep a running note of the modified area of the scratch layer.
' The compositor can use this information to only regenerate the compositor cache area that's changed since the
' last repaint event.  Note that the m_ModifiedRectF may be cleared between accesses, by design - you'll need to
' keep an eye on your usage of parameters in the GetModifiedUpdateRectF function.
'
'If you want the absolute modified area since the stroke began, you can use m_TotalModifiedRectF, which is not
' cleared until the current stroke is released.
Private m_UnionRectRequired As Boolean
Private m_ModifiedRectF As RectF, m_TotalModifiedRectF As RectF

'The number of mouse events in the *current* brush stroke.  This value is reset after every mouse release.
' The compositor uses this to know when to fully regenerate the paint cache from scratch.
Private m_NumOfMouseEvents As Long

'pd2D is used for certain brush styles
Private m_Painter As pd2DPainter, m_Surface As pd2DSurface

'To improve responsiveness, we measure the time delta between viewport refreshes.  If painting is happening fast enough,
' we coalesce screen updates together, as they are (by far) the most time-consuming segment of paint rendering;
' similarly, if painting is too slow, we temporarily reduce viewport update frequency until painting "catches up."
Private m_TimeSinceLastRender As Currency, m_NetTimeToRender As Currency, m_NumRenders As Long, m_FramesDropped As Long

Public Function GetBrushPreviewQuality() As PD_PerformanceSetting
    GetBrushPreviewQuality = m_BrushPreviewQuality
End Function

Public Function GetBrushPreviewQuality_GDIPlus() As GP_InterpolationMode
    If USE_PAINTBRUSH_DEBUG_QUALITIES Then
        If (m_BrushPreviewQuality = PD_PERF_FASTEST) Then
            GetBrushPreviewQuality_GDIPlus = GP_IM_NearestNeighbor
        ElseIf (m_BrushPreviewQuality = PD_PERF_BESTQUALITY) Then
            GetBrushPreviewQuality_GDIPlus = GP_IM_HighQualityBicubic
        Else
            GetBrushPreviewQuality_GDIPlus = GP_IM_Bilinear
        End If
    Else
        If (g_ViewportPerformance = PD_PERF_FASTEST) Then
            GetBrushPreviewQuality_GDIPlus = GP_IM_NearestNeighbor
        ElseIf (g_ViewportPerformance = PD_PERF_BESTQUALITY) Then
            GetBrushPreviewQuality_GDIPlus = GP_IM_HighQualityBicubic
        Else
            GetBrushPreviewQuality_GDIPlus = GP_IM_Bilinear
        End If
    End If
End Function

'Universal brush settings, applicable for most sources.  (I say "most" because some settings can contradict each other;
' for example, a "locked" alpha mode + "erase" blend mode makes little sense, but it is technically possible to set
' those values simultaneously.)
Public Function GetBrushAlphaMode() As PD_AlphaMode
    GetBrushAlphaMode = m_BrushAlphamode
End Function

Public Function GetBrushAntialiasing() As PD_2D_Antialiasing
    GetBrushAntialiasing = m_BrushAntialiasing
End Function

Public Function GetBrushBlendMode() As PD_BlendMode
    GetBrushBlendMode = m_BrushBlendmode
End Function

Public Function GetBrushFlow() As Single
    GetBrushFlow = m_BrushFlow
End Function

Public Function GetBrushHardness() As Single
    GetBrushHardness = m_BrushHardness
End Function

Public Function GetBrushOpacity() As Single
    GetBrushOpacity = m_BrushOpacity
End Function

Public Function GetBrushSize() As Single
    GetBrushSize = m_BrushSize
End Function

Public Function GetBrushSource() As PD_BrushSource
    GetBrushSource = m_BrushSource
End Function

Public Function GetBrushSourceColor() As Long
    GetBrushSourceColor = m_BrushSourceColor
End Function

Public Function GetBrushSpacing() As Single
    GetBrushSpacing = m_BrushSpacing
End Function

Public Function GetBrushStyle() As PD_BrushStyle
    GetBrushStyle = m_BrushStyle
End Function

'Property set functions.  Note that not all brush properties are used by all styles.
' (e.g. "brush hardness" is not used by "pencil" style brushes, etc)
Public Sub SetBrushAlphaMode(Optional ByVal newAlphaMode As PD_AlphaMode = LA_NORMAL)
    If (newAlphaMode <> m_BrushAlphamode) Then
        m_BrushAlphamode = newAlphaMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushAntialiasing(Optional ByVal newAntialiasing As PD_2D_Antialiasing = P2_AA_HighQuality)
    If (newAntialiasing <> m_BrushAntialiasing) Then
        m_BrushAntialiasing = newAntialiasing
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushBlendMode(Optional ByVal newBlendMode As PD_BlendMode = BL_NORMAL)
    If (newBlendMode <> m_BrushBlendmode) Then
        m_BrushBlendmode = newBlendMode
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushFlow(Optional ByVal newFlow As Single = 100#)
    If (newFlow <> m_BrushFlow) Then
        m_BrushFlow = newFlow
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushHardness(Optional ByVal newHardness As Single = 100#)
    newHardness = newHardness / 100
    If (newHardness <> m_BrushHardness) Then
        m_BrushHardness = newHardness
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushOpacity(ByVal newOpacity As Single)
    If (newOpacity <> m_BrushOpacity) Then
        m_BrushOpacity = newOpacity
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushPreviewQuality(ByVal newQuality As PD_PerformanceSetting)
    If (newQuality <> m_BrushPreviewQuality) Then
        m_BrushPreviewQuality = newQuality
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSize(ByVal newSize As Single)
    If (newSize <> m_BrushSize) Then
        m_BrushSize = newSize
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSource(ByVal newSource As PD_BrushSource)
    If (newSource <> m_BrushSource) Then
        m_BrushSource = newSource
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSourceColor(Optional ByVal newColor As Long = vbWhite)
    If (newColor <> m_BrushSourceColor) Then
        m_BrushSourceColor = newColor
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushSpacing(ByVal newSpacing As Single)
    newSpacing = newSpacing / 100
    If (newSpacing <> m_BrushSpacing) Then
        m_BrushSpacing = newSpacing
        m_BrushIsReady = False
    End If
End Sub

Public Sub SetBrushStyle(ByVal newStyle As PD_BrushStyle)
    If (newStyle <> m_BrushStyle) Then
        m_BrushStyle = newStyle
        m_BrushIsReady = False
    End If
End Sub

Public Function GetBrushProperty(ByVal bProperty As PD_BrushAttributes) As Variant
    
    Select Case bProperty
        Case BA_AlphaMode
            GetBrushProperty = GetBrushAlphaMode()
        Case BA_Antialiasing
            GetBrushProperty = GetBrushAntialiasing()
        Case BA_BlendMode
            GetBrushProperty = GetBrushBlendMode()
        Case BA_Flow
            GetBrushProperty = GetBrushFlow()
        Case BA_Hardness
            GetBrushProperty = GetBrushHardness()
        Case BA_Opacity
            GetBrushProperty = GetBrushOpacity()
        Case BA_Size
            GetBrushProperty = GetBrushSize()
        Case BA_Source
            GetBrushProperty = GetBrushSource()
        Case BA_SourceColor
            GetBrushProperty = GetBrushSourceColor()
        Case BA_Spacing
            GetBrushProperty = GetBrushSpacing()
        Case BA_Style
            GetBrushProperty = GetBrushStyle()
    End Select
    
End Function

Public Sub SetBrushProperty(ByVal bProperty As PD_BrushAttributes, ByVal newPropValue As Variant)
    
    Select Case bProperty
        Case BA_AlphaMode
            SetBrushAlphaMode newPropValue
        Case BA_Antialiasing
            SetBrushAntialiasing newPropValue
        Case BA_BlendMode
            SetBrushBlendMode newPropValue
        Case BA_Flow
            SetBrushFlow newPropValue
        Case BA_Hardness
            SetBrushHardness newPropValue
        Case BA_Opacity
            SetBrushOpacity newPropValue
        Case BA_Size
            SetBrushSize newPropValue
        Case BA_Source
            SetBrushSource newPropValue
        Case BA_SourceColor
            SetBrushSourceColor newPropValue
        Case BA_Spacing
            SetBrushSpacing newPropValue
        Case BA_Style
            SetBrushStyle newPropValue
    End Select
    
End Sub

Public Sub CreateCurrentBrush(Optional ByVal alsoCreateBrushOutline As Boolean = True, Optional ByVal forceCreation As Boolean = False)
        
    If ((Not m_BrushIsReady) Or forceCreation Or (Not m_BrushCreatedAtLeastOnce)) Then
    
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
    
        'At present, brush styles correspond nicely to brush engines.
        If (m_BrushStyle = BS_Pencil) Then
            m_BrushEngine = BE_GDIPlus
        ElseIf (m_BrushStyle = BS_SoftBrush) Then
            m_BrushEngine = BE_PhotoDemon
        End If
        
        Select Case m_BrushEngine
            
            Case BE_GDIPlus
                'For now, create a circular pen at the current size
                If (m_GDIPPen Is Nothing) Then Set m_GDIPPen = New pd2DPen
                Drawing2D.QuickCreateSolidPen m_GDIPPen, m_BrushSize, m_BrushSourceColor, , P2_LJ_Round, P2_LC_Round
                
            Case BE_PhotoDemon
                
                'Build a new brush reference image that reflects the current brush properties
                m_BrushSizeInt = Int(m_BrushSize + 0.999999)
                If (m_BrushStyle = BS_SoftBrush) Then CreateSoftBrushReference_PD
                m_SrcPenDIB.SetInitialAlphaPremultiplicationState True
                
                'We also need to calculate a brush spacing reference.  A spacing of 1 means that every pixel in
                ' the current stroke is dabbed.  From a performance perspective, this is simply not feasible for
                ' large brushes, so avoid it if possible.
                '
                'The "Automatic" setting (which maps to spacing = 0) automatically calculates spacing based on
                ' the current brush size.  (Basically, we dab every 1/2pi of a radius.)
                Dim tmpBrushSpacing As Single
                tmpBrushSpacing = m_BrushSize / PI_DOUBLE
                
                If (m_BrushSpacing > 0#) Then
                    tmpBrushSpacing = (m_BrushSpacing * tmpBrushSpacing)
                End If
                
                'The module-level spacing check is an integer (because we Mod it to test for paint dabs)
                m_BrushSpacingCheck = Int(tmpBrushSpacing + 0.5)
                If (m_BrushSpacingCheck < 1) Then m_BrushSpacingCheck = 1
                
                'Want to use some arbitrary DIB for testing purposes?  Uncomment the lines below.
                'Dim testImgPath As String
                'testImgPath = "C:\PhotoDemon v4\PhotoDemon\no_sync\Images from testers\brush_test_500.png"
                '
                'If (m_SrcPenDIB Is Nothing) Then Set m_SrcPenDIB = New pdDIB
                'Loading.QuickLoadImageToDIB testImgPath, m_SrcPenDIB, False, False, False
                'SetBrushSize m_SrcPenDIB.GetDIBWidth
                
                'Want to the GDI+ renderer (instead of GDI)?  Uncomment these two lines, then visit the
                ' ApplyPaintDab() function and uncomment the GDI+ renderer comment there.
                ' (This will be needed in the future for rotating and/or skewing the brush "on the fly"
                '  based on brush dynamics.)
                'If (m_CustomPenImage Is Nothing) Then Set m_CustomPenImage = New pd2DSurface
                'm_CustomPenImage.CreateSurfaceFromFile testImgPath
                
        End Select
        
        'Whenever we create a new brush, we should also refresh the current brush outline
        If alsoCreateBrushOutline Then CreateCurrentBrushOutline
        
        m_BrushIsReady = True
        m_BrushCreatedAtLeastOnce = True
        
        pdDebug.LogAction "Paintbrush.CreateCurrentBrush took " & VBHacks.GetTimeDiffNowAsString(startTime)
        
    End If
    
End Sub

Private Sub CreateSoftBrushReference_MyPaint()

    'Initialize our reference DIB as necessary
    If (m_SrcPenDIB Is Nothing) Then Set m_SrcPenDIB = New pdDIB
    If (m_SrcPenDIB.GetDIBWidth < m_BrushSizeInt - 1) Or (m_SrcPenDIB.GetDIBHeight < m_BrushSizeInt - 1) Then
        m_SrcPenDIB.CreateBlank m_BrushSizeInt, m_BrushSizeInt, 32, 0, 0
    Else
        m_SrcPenDIB.ResetDIB 0
    End If
    
    'Because we are only setting 255 possible different colors (one for each possible opacity, while the current
    ' color remains constant), this is a great candidate for lookup tables.  Note that for performance reasons,
    ' we're going to do something wacky, and prep our lookup table as *longs*.  This is (obviously) faster than
    ' setting each byte individually.
    Dim tmpR As Long, tmpG As Long, tmpB As Long
    tmpR = Colors.ExtractRed(m_BrushSourceColor)
    tmpG = Colors.ExtractGreen(m_BrushSourceColor)
    tmpB = Colors.ExtractBlue(m_BrushSourceColor)
    
    Dim cLookup() As Long
    ReDim cLookup(0 To 255) As Long
    
    Dim x As Long, y As Long, tmpMult As Single
    For x = 0 To 255
        tmpMult = CSng(x) / 255
        cLookup(x) = GDI_Plus.FillLongWithRGBA(tmpMult * tmpR, tmpMult * tmpG, tmpMult * tmpB, x)
    Next x
    
    'Prep manual per-pixel loop variables
    Dim dstImageData() As Long
    Dim tmpSA As SafeArray2D
    PrepSafeArray_Long tmpSA, m_SrcPenDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
    
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = m_SrcPenDIB.GetDIBWidth - 1
    finalY = m_SrcPenDIB.GetDIBHeight - 1
    
    'At present, we use a MyPaint-compatible system for calculating brush hardness.  This gives us comparable
    ' paint behavior against programs like MyPaint (obviously), Krita, and new versions of GIMP.
    ' Reference: https://github.com/mypaint/libmypaint/wiki/Using-Brushlib
    Dim brushAspectRatio As Single, brushAngle As Single
    
    'Some MyPaint-supported features are not currently exposed to the user.  Their hard-coded values appear below,
    ' and in the future, we may migrate these over to the UI.
    brushAspectRatio = 1#   '[1, #INF]
    brushAngle = 0#         '[0, 180] in degrees
    
    Dim refCos As Single, refSin As Single
    refCos = Cos(brushAngle / 360# * 2# * PI)
    refSin = Sin(brushAngle / 360# * 2# * PI)
    
    Dim dx As Single, dy As Single
    Dim dXr As Single, dYr As Single
    Dim brushRadius As Single, brushRadiusSquare As Single
    brushRadius = (m_BrushSize - 1#) / 2#
    brushRadiusSquare = brushRadius * brushRadius
    
    Dim dd As Single, pxOpacity As Single
    Dim brushHardness As Single
    brushHardness = m_BrushHardness
    If (brushHardness < 0.001) Then brushHardness = 0.001
    If (brushHardness > 0.999) Then brushHardness = 0.999
    
    'Loop through each pixel in the image, calculating per-pixel brush values as we go
    For x = initX To finalX
    For y = initY To finalY
    
        dx = x - brushRadius
        dy = y - brushRadius
        dXr = (dy * refSin + dx * refCos)
        dYr = (dy * refCos - dx * refSin) * brushAspectRatio
        
        dd = (dYr * dYr + dXr * dXr) / brushRadiusSquare
        
        If (dd > 1) Then
            pxOpacity = 0
        ElseIf (dd < brushHardness) Then
            pxOpacity = dd + 1 - (dd / brushHardness)
        Else
            pxOpacity = brushHardness / (1 - brushHardness) * (1 - dd)
        End If
        
        'NOTE: if you wanted to, you could apply flow here (e.g. pxOpacity * [0, 1])
        ' We ignore this for now as the MyPaint brush calculator isn't made available to the user.
        dstImageData(x, y) = cLookup(pxOpacity * 255)
        
        'TODO: optimize this function by only processing one quadrant, then mirroring the results to the
        ' other three matching quadrants.  (Obviously, this only works while aspect ratio = 1#)
        
    Next y
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4

End Sub

Private Sub CreateSoftBrushReference_PD()
    
    'Initialize our reference DIB as necessary
    If (m_SrcPenDIB Is Nothing) Then Set m_SrcPenDIB = New pdDIB
    If (m_SrcPenDIB.GetDIBWidth < m_BrushSizeInt) Or (m_SrcPenDIB.GetDIBHeight < m_BrushSizeInt) Then
        m_SrcPenDIB.CreateBlank m_BrushSizeInt, m_BrushSizeInt, 32, 0, 0
    Else
        m_SrcPenDIB.ResetDIB 0
    End If
    
    'Next, check for a few special cases.  First, brushes with maximum hardness don't need to be rendered manually.
    ' Instead, just plot an antialiased circle and call it good.
    Dim cSurface As pd2DSurface, cBrush As pd2DBrush
    If (m_BrushHardness = 1#) Then
        
        Drawing2D.QuickCreateSurfaceFromDC cSurface, m_SrcPenDIB.GetDIBDC, True
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        Drawing2D.QuickCreateSolidBrush cBrush, m_BrushSourceColor, m_BrushFlow
        m_Painter.FillCircleF cSurface, cBrush, m_BrushSize * 0.5, m_BrushSize * 0.5, m_BrushSize * 0.5
        
        Set cBrush = Nothing: Set cSurface = Nothing
    
    'If a brush has custom hardness, we're gonna have to render it manually.
    Else
        
        'Because we are only setting 255 possible different colors (one for each possible opacity, while the current
        ' color remains constant), this is a great candidate for lookup tables.  Note that for performance reasons,
        ' we're going to do something wacky, and prep our lookup table as *longs*.  This is (obviously) faster than
        ' setting each byte individually.
        Dim tmpR As Long, tmpG As Long, tmpB As Long
        tmpR = Colors.ExtractRed(m_BrushSourceColor)
        tmpG = Colors.ExtractGreen(m_BrushSourceColor)
        tmpB = Colors.ExtractBlue(m_BrushSourceColor)
        
        Dim cLookup() As Long
        ReDim cLookup(0 To 255) As Long
        
        'Calculate brush flow (which controls the opacity of individual dabs)
        Dim normMult As Single, flowMult As Single
        flowMult = m_BrushFlow * 0.01
        normMult = (1# / 255#) * flowMult
        
        Dim x As Long, y As Long, tmpMult As Single
        For x = 0 To 255
            tmpMult = CSng(x) * normMult
            cLookup(x) = GDI_Plus.FillLongWithRGBA(tmpMult * tmpR, tmpMult * tmpG, tmpMult * tmpB, x * flowMult)
        Next x
        
        'Next, we're going to do something weird.  If this brush is quite small, it's very difficult to plot subpixel
        ' data accurately.  Instead of messing with specialized calculations, we're just going to plot a larger
        ' temporary brush, then resample it down to the target size.  This is the least of many evils.
        Dim tmpBrushRequired As Boolean, tmpDIB As pdDIB
        Const BRUSH_SIZE_MIN_CUTOFF As Long = 15
        tmpBrushRequired = (m_BrushSize < BRUSH_SIZE_MIN_CUTOFF)
        
        'Prep manual per-pixel loop variables
        Dim dstImageData() As Long
        Dim tmpSA As SafeArray2D
        
        If tmpBrushRequired Then
            Set tmpDIB = New pdDIB
            tmpDIB.CreateBlank BRUSH_SIZE_MIN_CUTOFF, BRUSH_SIZE_MIN_CUTOFF, 32, 0, 0
            PrepSafeArray_Long tmpSA, tmpDIB
        Else
            PrepSafeArray_Long tmpSA, m_SrcPenDIB
        End If
        
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
        
        Dim initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        
        'For small brush sizes, we use the larger "temporary DIB" size as our target; the final result will be
        ' downsampled at the end.
        If tmpBrushRequired Then
            finalX = tmpDIB.GetDIBWidth - 1
            finalY = tmpDIB.GetDIBHeight - 1
        Else
            finalX = m_SrcPenDIB.GetDIBWidth - 1
            finalY = m_SrcPenDIB.GetDIBHeight - 1
        End If
        
        'After a good deal of testing, I've decided that I don't like the MyPaint system for calculating brush hardness.
        ' Their system behaves ridiculously at low "hardness" values, causing huge spacing issues for the brush.
        ' Instead, I'm using a system similar to PD's "vignette" tool, which yields much better results for beginners, IMO.
        Dim brushHardness As Single
        brushHardness = m_BrushHardness
        
        'Calculate interior and exterior brush radii.  Any pixels...
        ' - OUTSIDE the EXTERIOR radius are guaranteed to be fully transparent
        ' - INSIDE the INTERIOR radius are guaranteed to be fully opaque (or whatever the equivalent "max opacity" is for
        '    the current brush flow rate)
        ' - BETWEEN the exterior and interior radii will be feathered accordingly
        Dim brushRadius As Single, brushRadiusSquare As Single
        If tmpBrushRequired Then
            brushRadius = CSng(BRUSH_SIZE_MIN_CUTOFF) * 0.5
        Else
            brushRadius = m_BrushSize * 0.5
        End If
        brushRadiusSquare = brushRadius * brushRadius
        
        Dim innerRadius As Single, innerRadiusSquare As Single
        innerRadius = (brushRadius - 1) * (brushHardness * 0.99)
        innerRadiusSquare = innerRadius * innerRadius
        
        Dim radiusDifference As Single
        radiusDifference = (brushRadiusSquare - innerRadiusSquare)
        If (radiusDifference < 0.00001) Then radiusDifference = 0.00001
        radiusDifference = (1# / radiusDifference)
        
        Dim cx As Single, cy As Single
        Dim pxDistance As Single, pxOpacity As Single
        
        'Loop through each pixel in the image, calculating per-pixel brush values as we go
        For y = initY To finalY
        For x = initX To finalX
        
            'Calculate distance between this point and the idealized "center" of the brush
            cx = x - brushRadius
            cy = y - brushRadius
            pxDistance = (cx * cx + cy * cy)
            
            'Ignore pixels that lie outside the brush radius.  (These were initialized to full transparency,
            ' and we're simply gonna leave them that way.)
            If (pxDistance <= brushRadiusSquare) Then
                
                'If pixels lie *inside* the inner radius, set them to maximum opacity
                If (pxDistance <= innerRadiusSquare) Then
                    dstImageData(x, y) = cLookup(255)
                
                'If pixels lie somewhere between the inner radius and the brush radius, feather them appropriately
                Else
                
                    'Calculate the current distance as a linear amount between the inner radius (the smallest amount
                    ' of feathering this hardness value provides), and the outer radius (the actual brush radius)
                    pxOpacity = (brushRadiusSquare - pxDistance) * radiusDifference
                    
                    'Cube the result to produce a more gaussian-like fade
                    pxOpacity = pxOpacity * pxOpacity * pxOpacity
                    
                    'Pull the matching result from our lookup table
                    dstImageData(x, y) = cLookup(pxOpacity * 255#)
                    
                End If
                
            End If
        
        Next x
        Next y
        
        'Safely deallocate imageData()
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        
        'If a temporary brush was required (because the target brush is so small), downscale it to its
        ' final size now.
        If tmpBrushRequired Then
            GDI_Plus.GDIPlus_StretchBlt m_SrcPenDIB, 0#, 0#, m_BrushSize, m_BrushSize, tmpDIB, 0#, 0#, BRUSH_SIZE_MIN_CUTOFF, BRUSH_SIZE_MIN_CUTOFF, , GP_IM_HighQualityBilinear, , , True, True
        End If
        
    End If

End Sub

'As part of rendering the current brush, we also need to render a brush outline onto the canvas at the current
' mouse location.  The specific outline technique used varies by brush engine.
Public Sub CreateCurrentBrushOutline()
    
    Select Case m_BrushEngine
    
        'If this is a GDI+ brush, outline creation is pretty easy.  Assume a circular brush and simply
        ' create a path at that same size.  (Note that circles are defined by radius, while brushes are
        ' defined by diameter - hence the "/ 2".)
        Case BE_GDIPlus
        
            Set m_BrushOutlinePath = New pd2DPath
            
            'Single-pixel brushes are treated as a square for cursor purposes.
            If (m_BrushSize > 0#) Then
                If (m_BrushSize = 1) Then
                    m_BrushOutlinePath.AddRectangle_Absolute -0.75, -0.75, 0.75, 0.75
                Else
                    m_BrushOutlinePath.AddCircle 0, 0, m_BrushSize / 2 + 0.5
                End If
            End If
            
        'TODO!  Right now this is just a copy+paste of the GDI+ outline algorithm; we obviously need a more sophisticated
        ' one in the future.
        Case BE_PhotoDemon
            
            Set m_BrushOutlinePath = New pd2DPath
            
            'Single-pixel brushes are treated as a square for cursor purposes.
            If (m_BrushSize > 0#) Then
                If (m_BrushSize = 1) Then
                    m_BrushOutlinePath.AddRectangle_Absolute -0.75, -0.75, 0.75, 0.75
                Else
                    m_BrushOutlinePath.AddCircle 0, 0, m_BrushSize / 2 + 0.5
                End If
            End If
            
    End Select

End Sub

'Notify the brush engine of the current mouse position.  Coordinates should always be in *image* coordinate space,
' not screen space.  (Translation between spaces will be handled internally.)
Public Sub NotifyBrushXY(ByVal mouseButtonDown As Boolean, ByVal srcX As Single, ByVal srcY As Single, ByVal mouseTimeStamp As Long, ByRef srcCanvas As pdCanvas)
    
    Dim isFirstStroke As Boolean, isLastStroke As Boolean
    isFirstStroke = (Not m_MouseDown) And mouseButtonDown
    isLastStroke = m_MouseDown And (Not mouseButtonDown)
    
    'Perform a failsafe check for brush creation
    If (Not m_BrushIsReady) Then CreateCurrentBrush
    
    'If this is a MouseDown operation, we need to make sure the full paint engine is synchronized against any property
    ' changes that are applied "on-demand".
    If isFirstStroke Then
        
        'Switch the target canvas into high-resolution, non-auto-drop mode.  This basically means the mouse tracker
        ' reconstructs full mouse movement histories via GetMouseMovePointsEx, and it reports every last event to us,
        ' regardless of the delays involved.  (Normally, as mouse events become increasingly delayed, they are
        ' auto-dropped until the processor catches up.  We have other ways of working around that problem in the
        ' brush engine.)
        '
        'IMPORTANT NOTE: VirtualBox returns bad data via GetMouseMovePointsEx, so I now expose this setting to the user
        ' via the Tools > Options menu.  If the user disables high-res input, we will also ignore it.
        srcCanvas.SetMouseInput_HighRes Tools.GetToolSetting_HighResMouse()
        srcCanvas.SetMouseInput_AutoDrop False
        
        'Reset all internal mouse events trackers
        m_NumOfMouseEvents = 1
        m_NetTimeToRender = 0
        m_NumRenders = 0
        m_FramesDropped = 0
        
        'Make sure the current scratch layer is properly initialized
        Tools.InitializeToolsDependentOnImage
        pdImages(g_CurrentImage).ScratchLayer.SetLayerOpacity m_BrushOpacity
        pdImages(g_CurrentImage).ScratchLayer.SetLayerBlendMode m_BrushBlendmode
        pdImages(g_CurrentImage).ScratchLayer.SetLayerAlphaMode m_BrushAlphamode
        
        'Reset the "last mouse position" values to match the current ones
        m_MouseX = srcX
        m_MouseY = srcY
        
        'Notify the central "color history" manager of the color currently being used
        If (m_BrushSource = BS_Color) Then UserControls.PostPDMessage WM_PD_PRIMARY_COLOR_APPLIED, m_BrushSourceColor, , True
        
        'Initialize any relevant GDI+ objects for the current brush
        Drawing2D.QuickCreateSurfaceFromDC m_Surface, pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBDC, (m_BrushAntialiasing = P2_AA_HighQuality)
        
        'If we're directly using GDI+ for painting (by calling various GDI+ line commands), we need to explicitly set
        ' half-pixel offsets, so each pixel "coordinate" is treated as the *center* of the pixel instead of the top-left corner.
        ' (PD's paint engine handles this internally.)
        If (m_BrushEngine = BE_GDIPlus) Then m_Surface.SetSurfacePixelOffset P2_PO_Half
        
        'Reset any brush dynamics that are calculated on a per-stroke basis
        m_DistPixels = 0
        
    Else
        m_NumOfMouseEvents = m_NumOfMouseEvents + 1
    End If
    
    Dim startTime As Currency
    
    'If the mouse button is down, perform painting between the old and new points.
    ' (All painting occurs in image coordinate space, and is applied to the current image's scratch layer.)
    If mouseButtonDown Then
    
        'Want to profile this function?  Use this line of code (and the matching report line at the bottom of the function).
        VBHacks.GetHighResTime startTime
        
        'A separate function handles the actual rendering.
        ApplyPaintLine srcX, srcY, isFirstStroke
        
        'See if there are more points in the mouse move queue.  If there are, grab them all and stroke them immediately.
        Dim numPointsRemaining As Long
        numPointsRemaining = srcCanvas.GetNumMouseEventsPending
        
        If (numPointsRemaining > 0) And (Not isFirstStroke) Then
        
            Dim tmpMMP As MOUSEMOVEPOINT
            Dim imgX As Double, imgY As Double
            
            Do While srcCanvas.GetNextMouseMovePoint(VarPtr(tmpMMP))
                
                'The (x, y) points returned by this request are in the *hWnd's* coordinate space.  We must manually convert them
                ' to the image coordinate space.
                If Drawing.ConvertCanvasCoordsToImageCoords(srcCanvas, pdImages(g_CurrentImage), tmpMMP.x, tmpMMP.y, imgX, imgY) Then
                
                    'The paint layer is always full-size, so we don't need to perform a separate "image space to layer space"
                    ' coordinate conversion here.
                    ApplyPaintLine imgX, imgY, False
                    
                End If
                
            Loop
        
        End If
        
        'Notify the scratch layer of our updates
        pdImages(g_CurrentImage).ScratchLayer.NotifyOfDestructiveChanges
        
        'Report paint tool render times, as relevant
        'Debug.Print "Paint tool render timing: " & Format(CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000), "0000.00") & " ms"
    
    'The previous x/y coordinate trackers are updated automatically when the mouse is DOWN.  When the mouse is UP, we must manually
    ' modify those values.
    Else
        m_MouseX = srcX
        m_MouseY = srcY
    End If
    
    'With all painting tasks complete, update all old state values to match the new state values.
    m_MouseDown = mouseButtonDown
    
    'Unlike other drawing tools, the paintbrush engine controls viewport redraws.  This allows us to optimize behavior
    ' if we fall behind, and a long queue of drawing actions builds up.
    '
    '(Note that we only request manual redraws if the mouse is currently down; if the mouse *isn't* down, the canvas
    ' handles this for us.)
    If mouseButtonDown Then UpdateViewportWhilePainting isFirstStroke, startTime, srcCanvas
    
    'If the mouse button has been released, we can also release our internal GDI+ objects.
    ' (Note that the current *brush* resources are *not* released, by design.)
    If isLastStroke Then
        
        Set m_Surface = Nothing
        'm_MouseX = -1000000#
        'm_MouseY = -1000000#
        
        'Reset the target canvas's mouse handling behavior
        srcCanvas.SetMouseInput_HighRes False
        srcCanvas.SetMouseInput_AutoDrop True
        
    End If
    
End Sub

'While painting, we use a (fairly complicated) set of heuristics to decide when to update the primary viewport.
' We don't want to update it on every paint stroke event, as compositing the full viewport can be a very
' time-consuming process (especially for large images and/or images with many layers).
Private Sub UpdateViewportWhilePainting(ByVal isFirstStroke As Boolean, ByVal strokeStartTime As Currency, ByRef srcCanvas As pdCanvas)

    'If this is the first paint stroke, we always want to update the viewport to reflect that.
    Dim updateViewportNow As Boolean
    updateViewportNow = isFirstStroke
    
    'In the background, paint tool rendering is uncapped.  (60+ fps is achievable on most modern PCs, thankfully.)
    ' However, relaying those paint tool updates to the screen is a time-consuming process, as we have to composite
    ' the full image, apply color management, calculate zoom, and a whole bunch of other crap.  Because of this,
    ' it improves the user experience to run background paint calculations and on-screen viewport updates at
    ' different framerates, with an emphasis on making sure the *background* paint tool rendering gets top priority.
    If (Not updateViewportNow) Then
        
        'If this is the first frame we're rendering (which should have already been caught by the "isFirstStroke"
        ' check above), force a render
        If (m_NumRenders > 0) Then
        
            'Perform some quick heuristics to determine if brush performance is lagging; if it is, we can
            ' artificially delay viewport updates to compensate.  (On large images and/or at severe zoom-out values,
            ' viewport rendering consumes a disproportionate portion of the brush rendering process.)
            'Debug.Print "Average render time: " & Format$((m_NetTimeToRender / m_NumRenders) * 1000, "0000") & " ms"
            
            'Calculate an average per-frame render time for the current stroke, in ms.
            Dim avgFrameTime As Currency
            avgFrameTime = (m_NetTimeToRender / m_NumRenders) * 1000
            
            'If our average rendering time is "good" (above 15 fps), allow viewport updates to occur "in realtime",
            ' e.g. as fast as the background brush rendering.
            If (avgFrameTime < 66) Then
                updateViewportNow = True
            
            'If our average frame rendering time drops below 15 fps, start dropping viewport rendering frames, but only
            ' until we hit the (barely workable) threshold of 2 fps - at that point, we have to provide visual feedback,
            ' whatever the cost.
            Else
                
                'Never skip so many frames that viewport updates drop below 2 fps.  (This is absolutely a
                ' "worst-case" scenario, and it should never be relevant except on the lowliest of PCs.)
                updateViewportNow = (VBHacks.GetTimerDifferenceNow(m_TimeSinceLastRender) * 1000 > 500#)
                
                'If we're somewhere between 2 and 15 fps, keep an eye on how many frames we're dropping.  If we drop
                ' *too* many, the performance gain is outweighed by the obnoxiousness of stuttering screen renders.
                If (Not updateViewportNow) Then
                    
                    'This frame is a candidate for dropping.
                    Dim frameCutoff As Long
                    
                    'Next, determine how many frames we're allowed to drop.  As our average frame time increases,
                    ' we get more aggressive about dropping frames to compensate.  (This sliding scale tops out at
                    ' dropping 5 consecutive frames, which is pretty damn severe - but note that framerate drops
                    ' are also limited by the 2 fps check before this If/Then block.)
                    If (avgFrameTime < 100) Then
                        frameCutoff = 1
                    ElseIf (avgFrameTime < 133) Then
                        frameCutoff = 2
                    ElseIf (avgFrameTime < 167) Then
                        frameCutoff = 3
                    ElseIf (avgFrameTime < 200) Then
                        frameCutoff = 4
                    Else
                        frameCutoff = 5
                    End If
                    
                    'Keep track of how many frames we've dropped in a row
                    m_FramesDropped = m_FramesDropped + 1
                    
                    'If we've dropped too many frames proportionate to the current framerate, cancel this drop and
                    ' update the viewport.
                    If (m_FramesDropped > frameCutoff) Then updateViewportNow = True
                    
                End If
                
            End If
        
        End If
        
    End If
    
    'If a viewport update is required, composite the full layer stack prior to updating the screen
    If updateViewportNow Then
        
        'Reset the frame drop counter and the "time since last viewport render" tracker
        m_FramesDropped = 0
        VBHacks.GetHighResTime m_TimeSinceLastRender
        ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), srcCanvas, , pdImages(g_CurrentImage).GetActiveLayerIndex
    
    'If not enough time has passed since the last redraw, simply update the cursor
    Else
        ViewportEngine.Stage4_FlipBufferAndDrawUI pdImages(g_CurrentImage), srcCanvas
    End If
    
    'Update our running "time to render" tracker
    m_NetTimeToRender = m_NetTimeToRender + VBHacks.GetTimerDifferenceNow(strokeStartTime)
    m_NumRenders = m_NumRenders + 1
    
End Sub

'Formally render a line between the old mouse (x, y) coordinate pair and this new pair.  Replacement of the old (x, y) pair
' with the new coordinates is handled automatically.
Private Sub ApplyPaintLine(ByVal srcX As Single, ByVal srcY As Single, ByVal isFirstStroke As Boolean)
    
    'Calculate new modification rects, e.g. the portion of the paintbrush layer affected by this stroke.
    ' (The central compositor requires this information for its optimized paintbrush renderer.)
    UpdateModifiedRect srcX, srcY, isFirstStroke
    
    'Next, perform line rendering based on the current brush style.
    ' (At present, GDI+ is used to render a very basic brush.  More advanced styles are coming soon.)
    Select Case m_BrushEngine
            
        Case BE_GDIPlus
    
            'GDI+ refuses to draw a line if the start and end points match; this isn't documented (as far as I know),
            ' but it may exist to provide backwards compatibility with GDI, which deliberately leaves the last point
            ' of a line unplotted, in case you are drawing multiple connected lines.  Because of this, we have to
            ' manually render a dab at the initial starting position.
            If isFirstStroke Then
                m_Painter.DrawLineF m_Surface, m_GDIPPen, srcX, srcY, srcX - 0.1, srcY - 0.1
            Else
                m_Painter.DrawLineF m_Surface, m_GDIPPen, m_MouseX, m_MouseY, srcX, srcY
            End If
            
        Case BE_PhotoDemon
        
            'First strokes can just be applied as a single dab; this spares us attempting to calculate things like
            ' brush dynamics (which don't exist yet, as we have no point history).
            If isFirstStroke Then
                ApplyPaintDab srcX, srcY
            Else
                
                'If the target point is identical to the last point we rendered, ignore it (as the line between
                ' two identical points is "undefined", and not all line rasterizers detect this case successfully).
                If (srcX <> m_MouseX) Or (srcY <> m_MouseY) Then
                    ManuallyCalculateBrushPoints srcX, srcY
                End If
                
            End If
            
    End Select
    
    'Update the "old" mouse coordinate trackers
    m_MouseX = srcX
    m_MouseY = srcY
    
End Sub

'Calculate all point positions between (srcX, srcY) and the previous coordinates, and dab each point in turn.
' Note that I've implemented a number of different brush line algorithms; at present, a voxel-traversal algorithm
' is used instead of the more-obvious Bresenham method, as it provides proper sub-pixel coverage.
Private Sub ManuallyCalculateBrushPoints(ByVal srcX As Single, ByVal srcY As Single)
    
    'Want to use a traditional Bresenham rasterizer?  Here you go:
    'CalcPoints_Bresenham srcX, srcY
    
    'Voxel-traversal provides better support for floating-point brush sizes:
    CalcPoints_VoxelTraversal srcX, srcY
    
End Sub

Private Sub CalcPoints_VoxelTraversal(ByVal srcX As Single, ByVal srcY As Single)

    'TEST 3: voxel traversal approach based on "A Fast Voxel Traversal Algorithm for Ray Tracing."
    ' link: http://www.cse.yorku.ca/~amana/research/grid.pdf
    '
    'This is a highly efficient way to test every pixel "collision" against a line, by only testing pixel
    ' intersections.  There is a penalty at start-up (like most line algorithms), but traversal itself is
    ' extremely fast *and* friendly toward starting/ending floating-point coords.
    
    'Calculate directionality.  Note that I've manually added handling for the special case of horizontal
    ' and vertical lines.  (What I *haven't* yet implemented is speed-optimized versions of those special
    ' cases!)
    Dim stepX As Long, stepY As Long
    If (srcX > m_MouseX) Then
        stepX = 1
    Else
        If (srcX < m_MouseX) Then stepX = -1 Else stepX = 0
    End If
    If (srcY > m_MouseY) Then
        stepY = 1
    Else
        If (srcY < m_MouseY) Then stepY = -1 Else stepY = 0
    End If
    
    'Calculate deltas and termination conditions.  Note that these are all floating-point values, so we could
    ' theoretically support sub-pixel traversal conditions.  (At present, we only traverse full pixels.)
    Dim tDeltaX As Single, tMaxX As Single
    If (stepX <> 0) Then tDeltaX = PDMath.Min2Float_Single(CSng(stepX) / (srcX - m_MouseX), 10000000#) Else tDeltaX = 10000000#
    If (stepX > 0) Then tMaxX = tDeltaX * (1 - m_MouseX + Int(m_MouseX)) Else tMaxX = tDeltaX * (m_MouseX - Int(m_MouseX))
    
    Dim tDeltaY As Single, tMaxY As Single
    If (stepY <> 0) Then tDeltaY = PDMath.Min2Float_Single(CSng(stepY) / (srcY - m_MouseY), 10000000#) Else tDeltaY = 10000000#
    If (stepY > 0) Then tMaxY = tDeltaY * (1 - m_MouseY + Int(m_MouseY)) Else tMaxY = tDeltaY * (m_MouseY - Int(m_MouseY))
    
    'After some testing, I'm pretty pleased with the integer-only results of the traversal algorithm,
    ' so I've gone ahead and declared the traversal trackers as integer-only.  This doesn't do much for
    ' performance (as this algorithm is already highly optimized), but it does simplify some of our
    ' subsequent calculations.
    Dim x As Long, y As Long
    x = Int(m_MouseX)
    y = Int(m_MouseY)
    
    'Start plotting points.  Note that - by design, the first point is *not* manually rendered.
    Do
        
        'Apply this dab.
        ApplyPaintDab x, y
        
        'See if our next voxel (pixel) intersection occurs on a horizontal or vertical edge, and increase our
        ' running offset proportionally.
        If (tMaxX < tMaxY) Then
            tMaxX = tMaxX + tDeltaX
            x = x + stepX
        Else
            tMaxY = tMaxY + tDeltaY
            y = y + stepY
        End If
        
        'Check for traversal past the end of the destination voxel
        If (tMaxX > 1) Then
            If (tMaxY > 1) Then Exit Do
        End If
        
    Loop
    
End Sub

'Bresenham line rasterizer.  Currently unused, but provided for educational purposes.
Private Sub CalcPoints_Bresenham(ByVal srcX As Single, ByVal srcY As Single)

    'This is a barebones Bresenham implementation.  It would be difficult to improve speed much beyond
    ' this code, short of specialized per-brush implementations, so this is a nice baseline for "fast but
    ' sketchy pixel coverage."  (Note that performance of this function itself is irrelevant -- the cost of
    ' stroke rendering lies entirely in rendering the brush itself.)
    
    'Like any Bresenham implementation, all calculations are done as integers
    Dim x0 As Long, x1 As Long, y0 As Long, y1 As Long
    x0 = m_MouseX
    y0 = m_MouseY
    x1 = srcX
    y1 = srcY
    
    'Calculate deltas
    Dim dx As Long, dy As Long
    dx = Abs(x1 - x0)
    dy = Abs(y1 - y0)
    
    'Calculate step directionality.
    ' (NOTE: this function does not currently implement specialized detection for horizontal or vertical lines.)
    Dim sX As Long, sY As Long
    If (x0 < x1) Then sX = 1 Else sX = -1
    If (y0 < y1) Then sY = 1 Else sY = -1
    
    'Running "errors" are used to bump the running pixel calculations in x or y directions
    Dim runningErr As Long, e2 As Long
    runningErr = dx - dy
    
    Do
        
        'Once we hit the final pixel, exit immediately.
        If ((x0 = x1) And (y0 = y1)) Then
            Exit Do
        End If
        
        'Calculate a new error, and determine if we need to advance in the X or Y direction
        e2 = 2 * runningErr
        If (e2 > -dy) Then
            runningErr = runningErr - dy
            x0 = x0 + sX
        End If
        
        If (e2 < dx) Then
            runningErr = runningErr + dx
            y0 = y0 + sY
        End If
        
        'Dab the target pixel
        ApplyPaintDab x0, y0
        
    Loop
    
End Sub

'Apply a single paint dab to the target position.  Note that dab opacity is currently hard-coded at 100%; flow is controlled
' at brush creation time (instead of on-the-fly).  This may change depending on future brush dynamics implementations.
Private Sub ApplyPaintDab(ByVal srcX As Single, ByVal srcY As Single, Optional ByVal dabOpacity As Single = 1#)
    
    Dim allowedToDab As Boolean: allowedToDab = True
    
    'If brush dynamics are active, we only dab the brush if certain criteria are met.  (For example, if enough pixels have
    ' elapsed since the last dab, as controlled by the Brush Spacing parameter.)
    If (m_BrushSpacingCheck > 1) Then allowedToDab = ((m_DistPixels Mod m_BrushSpacingCheck) = 0)
    
    If allowedToDab Then
        
        'TODO: certain features (like brush rotation) will require a GDI+ surface.  Simple brushes can use GDI's AlphaBlend
        ' for a performance boost, however.
        m_SrcPenDIB.AlphaBlendToDCEx pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBDC, Int(srcX - m_BrushSize \ 2), Int(srcY - m_BrushSize \ 2), Int(m_BrushSize), Int(m_BrushSize), 0, 0, Int(m_BrushSize), Int(m_BrushSize), dabOpacity * 255
        'm_Painter.DrawSurfaceF m_Surface, srcX - m_BrushSize / 2, srcY - m_BrushSize / 2, m_CustomPenImage, dabOpacity * 100
        
    End If
    
    'Each time we make a new dab, we keep a running tally of how many pixels we've traversed.  Some brush dynamics (e.g. spacing)
    ' rely on this value for correct rendering behavior.
    m_DistPixels = m_DistPixels + 1
    
End Sub

'Whenever we receive notifications of a new mouse (x, y) pair, you need to call this sub to calculate a new "affected area" rect.
' The compositor uses this "affected area" rect to minimize the amount of rendering work it needs to perform.
Private Sub UpdateModifiedRect(ByVal newX As Single, ByVal newY As Single, ByVal isFirstStroke As Boolean)

    'Start by calculating the affected rect for just this stroke.
    Dim tmpRectF As RectF
    If (newX < m_MouseX) Then
        tmpRectF.Left = newX
        tmpRectF.Width = m_MouseX - newX
    Else
        tmpRectF.Left = m_MouseX
        tmpRectF.Width = newX - m_MouseX
    End If
    
    If (newY < m_MouseY) Then
        tmpRectF.Top = newY
        tmpRectF.Height = m_MouseY - newY
    Else
        tmpRectF.Top = m_MouseY
        tmpRectF.Height = newY - m_MouseY
    End If
    
    'Inflate the rect calculation by the size of the current brush, while accounting for the possibility of antialiasing
    ' (which may extend up to 1.0 pixel outside the calculated boundary area).
    Dim halfBrushSize As Single
    halfBrushSize = m_BrushSize / 2 + 1#
    
    tmpRectF.Left = tmpRectF.Left - halfBrushSize
    tmpRectF.Top = tmpRectF.Top - halfBrushSize
    
    halfBrushSize = halfBrushSize * 2
    tmpRectF.Width = tmpRectF.Width + halfBrushSize
    tmpRectF.Height = tmpRectF.Height + halfBrushSize
    
    Dim tmpOldRectF As RectF
    
    'If this is *not* the first modified rect calculation, union this rect with our previous update rect
    If m_UnionRectRequired And (Not isFirstStroke) Then
        tmpOldRectF = m_ModifiedRectF
        PDMath.UnionRectF m_ModifiedRectF, tmpRectF, tmpOldRectF
    Else
        m_UnionRectRequired = True
        m_ModifiedRectF = tmpRectF
    End If
    
    'Always calculate a running "total combined RectF", for use in the final merge step
    If isFirstStroke Then
        m_TotalModifiedRectF = tmpRectF
    Else
        tmpOldRectF = m_TotalModifiedRectF
        PDMath.UnionRectF m_TotalModifiedRectF, tmpRectF, tmpOldRectF
    End If
    
End Sub

'Return the area of the image modified by the current stroke.  By default, the running modified rect is erased after a call to
' this function, but this behavior can be toggled by resetRectAfter.  Also, if you want to get the full modified rect since this
' paint stroke began, you can set the GetModifiedRectSinceStrokeBegan parameter to TRUE.  Note that when
' GetModifiedRectSinceStrokeBegan is TRUE, the resetRectAfter parameter is ignored.
Public Function GetModifiedUpdateRectF(Optional ByVal resetRectAfter As Boolean = True, Optional ByVal GetModifiedRectSinceStrokeBegan As Boolean = False) As RectF
    If GetModifiedRectSinceStrokeBegan Then
        GetModifiedUpdateRectF = m_TotalModifiedRectF
    Else
        GetModifiedUpdateRectF = m_ModifiedRectF
        If resetRectAfter Then m_UnionRectRequired = False
    End If
End Function

Public Function GetNumOfStrokes() As Long
    GetNumOfStrokes = m_NumOfMouseEvents
End Function

'Want to commit your current brush work?  Call this function to make the brush results permanent.
Public Sub CommitBrushResults()
    
    'Reset the current mouse event counter
    m_NumOfMouseEvents = 0
    
    'Make a local copy of the paintbrush's bounding rect, and clip it to the layer's boundaries
    Dim tmpRectF As RectF
    tmpRectF = m_TotalModifiedRectF
    
    With tmpRectF
        If (.Left < 0) Then .Left = 0
        If (.Top < 0) Then .Top = 0
        If (.Width > pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBWidth) Then .Width = pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBWidth
        If (.Height > pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBHeight) Then .Height = pdImages(g_CurrentImage).ScratchLayer.layerDIB.GetDIBHeight
    End With
    
    'Committing brush results is actually pretty easy!
    
    'First, if the layer beneath the paint stroke is a raster layer, we simply want to merge the scratch
    ' layer onto it.
    If pdImages(g_CurrentImage).GetActiveLayer.IsLayerRaster Then
        
        Dim bottomLayerFullSize As Boolean
        With pdImages(g_CurrentImage).GetActiveLayer
            bottomLayerFullSize = ((.GetLayerOffsetX = 0) And (.GetLayerOffsetY = 0) And (.layerDIB.GetDIBWidth = pdImages(g_CurrentImage).Width) And (.layerDIB.GetDIBHeight = pdImages(g_CurrentImage).Height))
        End With
        
        pdImages(g_CurrentImage).MergeTwoLayers pdImages(g_CurrentImage).ScratchLayer, pdImages(g_CurrentImage).GetActiveLayer, bottomLayerFullSize, True, VarPtr(tmpRectF)
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_Layer, pdImages(g_CurrentImage).GetActiveLayerIndex
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Paint stroke", , , UNDO_Layer, g_CurrentTool
        
        'Reset the scratch layer
        pdImages(g_CurrentImage).ScratchLayer.layerDIB.ResetDIB 0
    
    'If the layer beneath this one is *not* a raster layer, let's add the stroke as a new layer, instead.
    Else
        
        'Before creating the new layer, check for an active selection.  If one exists, we need to preprocess
        ' the paint layer against it.
        If pdImages(g_CurrentImage).IsSelectionActive Then
            
            'A selection is active.  Pre-mask the paint scratch layer against it.
            Dim cBlender As pdPixelBlender
            Set cBlender = New pdPixelBlender
            cBlender.ApplyMaskToTopDIB pdImages(g_CurrentImage).ScratchLayer.layerDIB, pdImages(g_CurrentImage).MainSelection.GetMaskDIB, VarPtr(tmpRectF)
            
        End If
        
        Dim newLayerID As Long
        newLayerID = pdImages(g_CurrentImage).CreateBlankLayer(pdImages(g_CurrentImage).GetActiveLayerIndex)
        
        'Point the new layer index at our scratch layer
        pdImages(g_CurrentImage).PointLayerAtNewObject newLayerID, pdImages(g_CurrentImage).ScratchLayer
        pdImages(g_CurrentImage).GetLayerByID(newLayerID).SetLayerName g_Language.TranslateMessage("Paint layer")
        Set pdImages(g_CurrentImage).ScratchLayer = Nothing
        
        'Activate the new layer
        pdImages(g_CurrentImage).SetActiveLayerByID newLayerID
        
        'Notify the parent image of the new layer
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_Image_VectorSafe
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Paint stroke", , , UNDO_Image_VectorSafe, g_CurrentTool
        
        'Create a new scratch layer
        Tools.InitializeToolsDependentOnImage
        
    End If
    
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
    onScreenSize = Drawing.ConvertImageSizeToCanvasSize(m_BrushSize, pdImages(g_CurrentImage))
    
    Dim brushTooSmall As Boolean
    brushTooSmall = (onScreenSize < 7#)
    
    'Borrow a pair of UI pens from the main rendering module
    Dim innerPen As pd2DPen, outerPen As pd2DPen
    Drawing.BorrowCachedUIPens outerPen, innerPen
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetCanvas.hDC, True
    
    'Paint a target cursor - but *only* if the mouse is not currently down!
    Dim crossLength As Single, outerCrossBorder As Single
    crossLength = 3#
    outerCrossBorder = 0.5
    
    If (Not m_MouseDown) Then
        m_Painter.DrawLineF cSurface, outerPen, cursX, cursY - crossLength - outerCrossBorder, cursX, cursY + crossLength + outerCrossBorder
        m_Painter.DrawLineF cSurface, outerPen, cursX - crossLength - outerCrossBorder, cursY, cursX + crossLength + outerCrossBorder, cursY
        m_Painter.DrawLineF cSurface, innerPen, cursX, cursY - crossLength, cursX, cursY + crossLength
        m_Painter.DrawLineF cSurface, innerPen, cursX - crossLength, cursY, cursX + crossLength, cursY
    End If
    
    'If size allows, render a transformed brush outline onto the canvas as well
    If (Not brushTooSmall) Then
        
        'Get a copy of the current brush outline, transformed into position
        Dim copyOfBrushOutline As pd2DPath
        Set copyOfBrushOutline = New pd2DPath
        
        copyOfBrushOutline.CloneExistingPath m_BrushOutlinePath
        copyOfBrushOutline.ApplyTransformation canvasMatrix
        m_Painter.DrawPath cSurface, outerPen, copyOfBrushOutline
        m_Painter.DrawPath cSurface, innerPen, copyOfBrushOutline
        
    End If
    
    Set cSurface = Nothing
    
End Sub

'A brush is considered active if the mouse state is currently DOWN, or if it is up but we are still rendering a
' previous stroke.
Public Function IsBrushActive() As Boolean
    IsBrushActive = m_MouseDown
End Function

'Any specialized initialization tasks can be handled here.  This function is called early in the PD load process.
Public Sub InitializeBrushEngine()
    m_BrushPreviewQuality = PD_PERF_BALANCED
    m_BrushAntialiasing = P2_AA_HighQuality
    Drawing2D.QuickCreatePainter m_Painter
    m_MouseX = -1000000#
    m_MouseY = -1000000#
    m_BrushIsReady = False
    m_BrushCreatedAtLeastOnce = False
End Sub

'Before PD closes, you *must* call this function!  It will free any lingering brush resources (which are cached
' for performance reasons).
Public Sub FreeBrushResources()
    Set m_GDIPPen = Nothing
    Set m_CustomPenImage = Nothing
    Set m_BrushOutlineImage = Nothing
    Set m_BrushOutlinePath = Nothing
    Set m_Painter = Nothing
End Sub
