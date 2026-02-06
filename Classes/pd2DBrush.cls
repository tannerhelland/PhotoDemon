VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Graphics Fill ("Brush" in GDI/GDI+ parlance) Class
'Copyright 2014-2026 by Tanner Helland
'Created: 30/June/15 (but assembled from many bits written earlier)
'Last updated: 13/April/22
'Last update: fix the way global opacity affects pattern brush opacity
'
'This class manages a single brush (fill) instance.  Brushes are used to fill shapes, regions, and/or paths.
'
'At present, this class is primarily based on the capabilities of GDI+.  This may change going forward, but because
' GDI+ provides a nice baseline feature set, that's where I started.
'
'IMPORTANT NOTE!  Some property changes require us to destroy the current brush and create a new one from scratch.
' For example, a solid fill brush can change its color without creating a new handle, but a hatch brush cannot.
' (This is a built-in limitation of the GDI+ flat API.)  For that reason, you should not cache brush handles returned
' by this class.  Instead, you should always use the .GetHandle() function, as it will silently create new handles
' when necessary.
'
'This class also supports gradient brushes.  Gradient brushes are managed differently; they rely on a pd2DGradient class,
' which manages all gradient-related settings and simply passes this class a bare handle as necessary.
'
'Texture brushes are still TODO.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'GDI+ declares
Private Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal brushHatchStyle As GP_PatternStyle, ByVal brushForegroundRGBA As Long, ByVal brushBackgroundRGBA As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal srcColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal hImage As Long, ByVal textureWrapMode As GP_WrapMode, ByRef dstTexture As Long) As GP_Result
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As GP_Result
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GP_Result
'Private Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByRef dstColor As RGBQuad) As GP_Result
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByVal newColor As Long) As GP_Result
Private Declare Function GdipSetTextureTransform Lib "gdiplus" (ByVal hBrush As Long, ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result

'Current brush mode (solid fill, gradient, texture, etc)
Private m_BrushMode As PD_2D_BrushMode

'Solid brush settings
Private m_BrushRGBA As RGBQuad    'GDI+ requires RGB quads; we translate between this format as necessary

'Pattern brush settings
Private m_PatternStyle As PD_2D_PatternStyle
Private m_PatternRGBA1 As RGBQuad
Private m_PatternRGBA2 As RGBQuad

'Gradient brushes are more complicated, as they rely on a variable number of settings.
' We use a specialized class to manage those settings.
Private m_Gradient As pd2DGradient

'Gradient fills also require a boundary rect to define the start/end positions of the gradient line.
' You *must set this prior to requesting a gradient brush* or the gradient won't paint to the correct location!
Private m_BoundaryRect As RectF

'Texture brushes are much simpler if we maintain a copy of the source texture (e.g. this avoids messy cases like the caller
' creating a texture brush, then erasing the source image prior to freeing *this* brush).  For simplicity, the brush is stored
' inside a pdDIB object, with any clipping applied *prior* to creating the DIB.
Private m_TextureSrc As pdDIB
Private m_TextureWrapMode As PD_2D_WrapMode

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdSerialize

'Once a brush has been created, this handle value will be non-zero
Private m_BrushHandle As Long

'NOTE: texture brush settings are still TBD

'Get/set individual settings.
Friend Function GetBrushMode() As PD_2D_BrushMode
    GetBrushMode = m_BrushMode
End Function

Friend Function GetBrushColor() As Long
    GetBrushColor = RGB(m_BrushRGBA.Red, m_BrushRGBA.Green, m_BrushRGBA.Blue)
End Function

Friend Function GetBrushColorRGBA() As RGBQuad
    GetBrushColorRGBA = m_BrushRGBA
End Function

Friend Function GetBrushOpacity() As Single
    GetBrushOpacity = CSng(m_BrushRGBA.Alpha) / 2.55!
End Function

Friend Function GetBrushPatternStyle() As PD_2D_PatternStyle
    GetBrushPatternStyle = m_PatternStyle
End Function

Friend Function GetBrushPattern1Color() As Long
    GetBrushPattern1Color = RGB(m_PatternRGBA1.Red, m_PatternRGBA1.Green, m_PatternRGBA1.Blue)
End Function

Friend Function GetBrushPattern1Opacity() As Single
    GetBrushPattern1Opacity = CSng(m_PatternRGBA1.Alpha) / 2.55!
End Function

Friend Function GetBrushPattern2Color() As Long
    GetBrushPattern2Color = RGB(m_PatternRGBA2.Red, m_PatternRGBA2.Green, m_PatternRGBA2.Blue)
End Function

Friend Function GetBrushPattern2Opacity() As Single
    GetBrushPattern2Opacity = CSng(m_PatternRGBA2.Alpha) / 2.55!
End Function

Friend Function GetBrushGradientAllSettings() As String
    If (Not m_Gradient Is Nothing) Then GetBrushGradientAllSettings = m_Gradient.GetGradientAsString()
End Function
    
Friend Function GetBrushGradientShape() As PD_2D_GradientShape
    If (Not m_Gradient Is Nothing) Then GetBrushGradientShape = m_Gradient.GetGradientShape()
End Function

Friend Function GetBrushGradientAngle() As Single
    If (Not m_Gradient Is Nothing) Then GetBrushGradientAngle = m_Gradient.GetGradientAngle()
End Function

Friend Function GetBrushGradientWrapMode() As PD_2D_WrapMode
    If (Not m_Gradient Is Nothing) Then GetBrushGradientWrapMode = m_Gradient.GetGradientWrapMode()
End Function

Friend Function GetBrushGradientNodes() As String
    If (Not m_Gradient Is Nothing) Then GetBrushGradientNodes = m_Gradient.GetGradientNodes()
End Function

Friend Function GetBrushTextureWrapMode() As PD_2D_WrapMode
    GetBrushTextureWrapMode = m_TextureWrapMode
End Function

'Brush mode is an internal pd2D setting; as such, the entire brush needs to be re-created after changing this setting.
Friend Sub SetBrushMode(ByVal newSetting As PD_2D_BrushMode)
    If (newSetting <> m_BrushMode) Then
        m_BrushMode = newSetting
        If (m_BrushHandle <> 0) Then Me.ReleaseBrush
    End If
End Sub

Friend Function SetBrushColor(ByVal newSetting As Long) As Boolean
    
    'This function does *not* currently support OLE color constants; you need to manually uncomment
    ' the OleTranslateColor call, below, if you care about this.  (PD does not, and it's faster to
    ' skip the function call).  Also, you should read the MSDN docs for OleTranslateColor before
    ' enabling this line, as its return may be meaningful.
    'If OleTranslateColor(newSetting, 0, newSetting) Then 'failure - do what you want!
    
    'Extract RGB components
    m_BrushRGBA.Red = Colors.ExtractRed(newSetting)
    m_BrushRGBA.Green = Colors.ExtractGreen(newSetting)
    m_BrushRGBA.Blue = Colors.ExtractBlue(newSetting)
    
    'Apply immediately as necessary
    SetBrushColor = True
    If (m_BrushHandle <> 0) And (m_BrushMode = P2_BM_Solid) Then
        
        'Workaround for ByVal UDTs
        Dim tmpLong As Long
        GetMem4 VarPtr(m_BrushRGBA), tmpLong
        SetBrushColor = (GdipSetSolidFillColor(m_BrushHandle, tmpLong) = GP_OK)
        
    End If
    
End Function

'Color does *not* currently support OLE color constants; you need to manually add a call to
' OleTranslateColor if you want to blindly use system color constants
Friend Function SetBrushColorRGBA(ByRef newSetting As RGBQuad) As Boolean
    
    'If the current brush is *not* a solid-fill brush, we only need to reset it if
    ' the alpha-value changes.  (RGB changes do not affect gradient brushes, for example.)
    Dim origOpacity As Byte
    origOpacity = m_BrushRGBA.Alpha
    
    m_BrushRGBA = newSetting
    
    'Apply immediately as necessary
    SetBrushColorRGBA = True
    If (m_BrushHandle <> 0) Then
        
        If (m_BrushMode = P2_BM_Solid) Then
        
            'Workaround for ByVal UDTs
            Dim tmpLong As Long
            GetMem4 VarPtr(m_BrushRGBA), tmpLong
            SetBrushColorRGBA = (GdipSetSolidFillColor(m_BrushHandle, tmpLong) = GP_OK)
            
        Else
            If (origOpacity <> newSetting.Alpha) Then Me.ReleaseBrush
        End If
        
    End If
    
End Function

Friend Function SetBrushOpacity(ByVal newSetting As Single) As Boolean
    
    Dim origOpacity As Byte
    origOpacity = m_BrushRGBA.Alpha
    
    m_BrushRGBA.Alpha = Int(newSetting * 2.55! + 0.5!)
    
    'Apply immediately as necessary
    SetBrushOpacity = True
    If (m_BrushHandle <> 0) Then
        
        If (m_BrushMode = P2_BM_Solid) Then
            
            'Workaround for ByVal UDTs
            Dim tmpLong As Long
            GetMem4 VarPtr(m_BrushRGBA), tmpLong
            SetBrushOpacity = (GdipSetSolidFillColor(m_BrushHandle, tmpLong) = GP_OK)
        
        'For non-solid brushes, the entire brush must be recreated
        Else
            If (origOpacity <> m_BrushRGBA.Alpha) Then Me.ReleaseBrush
        End If
            
    End If
    
End Function

'GDI+ does not allow changes to hatch-brush settings.  Instead, the brush must be
' destroyed and re-created.
Friend Sub SetBrushPatternStyle(ByVal newSetting As PD_2D_PatternStyle)
    m_PatternStyle = newSetting
    If (m_BrushHandle <> 0) And (m_BrushMode = P2_BM_Pattern) Then Me.ReleaseBrush
End Sub

Friend Sub SetBrushPattern1Color(ByVal newSetting As Long)
    
    'This function does *not* currently support OLE color constants; you need to manually uncomment
    ' the OleTranslateColor call, below, if you care about this.  (PD does not, and it's faster to
    ' skip the function call).  Also, you should read the MSDN docs for OleTranslateColor before
    ' enabling this line, as its return may be meaningful.
    'If OleTranslateColor(newSetting, 0, newSetting) Then 'failure - do what you want!
    
    'Extract RGB components
    m_PatternRGBA1.Red = Colors.ExtractRed(newSetting)
    m_PatternRGBA1.Green = Colors.ExtractGreen(newSetting)
    m_PatternRGBA1.Blue = Colors.ExtractBlue(newSetting)
    
    'If the brush exists, it needs to be destroyed and recreated.  Since the user may be setting
    ' multiple properties at once, it's faster to just kill the brush; it will be re-created when
    ' its handle is next requested.
    If (m_BrushHandle <> 0) And (m_BrushMode = P2_BM_Pattern) Then Me.ReleaseBrush
    
End Sub

Friend Sub SetBrushPattern1Opacity(ByVal newSetting As Single)
    m_PatternRGBA1.Alpha = Int(newSetting * 2.55! + 0.5!)
    If (m_BrushHandle <> 0) And (m_BrushMode = P2_BM_Pattern) Then Me.ReleaseBrush
End Sub

Friend Sub SetBrushPattern2Color(ByVal newSetting As Long)
    
    'This function does *not* currently support OLE color constants; you need to manually uncomment
    ' the OleTranslateColor call, below, if you care about this.  (PD does not, and it's faster to
    ' skip the function call).  Also, you should read the MSDN docs for OleTranslateColor before
    ' enabling this line, as its return may be meaningful.
    'If OleTranslateColor(newSetting, 0, newSetting) Then 'failure - do what you want!
    
    'Extract RGB components
    m_PatternRGBA2.Red = Colors.ExtractRed(newSetting)
    m_PatternRGBA2.Green = Colors.ExtractGreen(newSetting)
    m_PatternRGBA2.Blue = Colors.ExtractBlue(newSetting)
    
    'If the brush exists, it needs to be destroyed and recreated.  Since the user may be setting
    ' multiple properties at once, it's faster to just kill the brush; it will be re-created when
    ' its handle is next requested.
    If (m_BrushHandle <> 0) And (m_BrushMode = P2_BM_Pattern) Then Me.ReleaseBrush
    
End Sub

Friend Sub SetBrushPattern2Opacity(ByVal newSetting As Single)
    m_PatternRGBA2.Alpha = Int(newSetting * 2.55! + 0.5!)
    If (m_BrushHandle <> 0) And (m_BrushMode = P2_BM_Pattern) Then Me.ReleaseBrush
End Sub

Friend Sub SetBrushGradientAllSettings(ByRef newSetting As String)
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.CreateGradientFromString newSetting
    'We don't need to relay gradient settings to GDI+; the gradient class handles it internally
End Sub
    
Friend Sub SetBrushGradientShape(ByVal newSetting As PD_2D_GradientShape)
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientShape newSetting
    'We don't need to relay gradient settings to GDI+; the gradient class handles it internally
End Sub

Friend Sub SetBrushGradientAngle(ByVal newSetting As Single)
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientAngle newSetting
    'We don't need to relay gradient settings to GDI+; the gradient class handles it internally
End Sub

Friend Sub SetBrushGradientWrapMode(ByVal newSetting As PD_2D_WrapMode)
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientWrapMode newSetting
    'We don't need to relay gradient settings to GDI+; the gradient class handles it internally
End Sub

Friend Sub SetBrushGradientNodes(ByRef newSetting As String)
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientNodes newSetting
    'We don't need to relay gradient settings to GDI+; the gradient class handles it internally
End Sub

'Cache the passed texture
Friend Sub SetBrushTextureFromDC(ByVal srcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcColorDepth As Long = 32)
    If (m_TextureSrc Is Nothing) Then Set m_TextureSrc = New pdDIB
    m_TextureSrc.CreateFromDC srcDC, srcX, srcY, srcWidth, srcHeight, srcColorDepth
End Sub

'Cache the passed texture (if requested).  The useSoftReference parameter will simply use your existing
' DIB as-is instead of making a local copy.  This cuts memory usage but you *MUST* keep your source DIB
' alive for the duration of this brush object.
Friend Sub SetBrushTextureFromDIB(ByRef srcDIB As pdDIB, Optional ByVal useSoftReference As Boolean = True)
    If useSoftReference Then
        Set m_TextureSrc = srcDIB
    Else
        If (m_TextureSrc Is Nothing) Then Set m_TextureSrc = New pdDIB
        m_TextureSrc.CreateFromExistingDIB srcDIB
    End If
End Sub

'Unlike other brush functions, this one requires that the brush has already been created.
' (This is due to the way the transform handle is cached, and our inability to ensure the lifetime
' of the source transform object.)
Friend Function SetBrushTextureTransform(ByRef srcTransform As pd2DTransform) As Boolean
    
    'Warn the caller that setting a transform on a null-brush isn't technically allowed,
    ' but attempt to work around that by creating the brush for them anyway
    If (m_BrushHandle = 0) Then
        InternalError "SetBrushTextureTransform", "Brush must exist before setting transform; creating now..."
        Me.CreateBrush
    End If
    
    'A brush handle *must* exist to proceed further
    If (m_BrushHandle <> 0) Then
        SetBrushTextureTransform = (GdipSetTextureTransform(m_BrushHandle, srcTransform.GetHandle) = GP_OK)
        If (Not SetBrushTextureTransform) Then InternalError "SetBrushTextureTransform", "GDI+ failure"
    Else
        InternalError "SetBrushTextureTransform", "null hBrush"
    End If
    
End Function

Friend Function SetBrushTextureWrapMode(ByVal newSetting As PD_2D_WrapMode) As Boolean
    m_TextureWrapMode = newSetting
    If (m_BrushHandle <> 0) Then
        SetBrushTextureWrapMode = (GdipSetTextureWrapMode(m_BrushHandle, newSetting) = GP_OK)
        If (Not SetBrushTextureWrapMode) Then InternalError "SetBrushTextureTransform", "GDI+ failure"
    Else
        SetBrushTextureWrapMode = True
    End If
End Function

Friend Sub SetBoundaryRect(ByRef srcRect As RectF)
    m_BoundaryRect = srcRect
    If (m_BrushMode = P2_BM_Gradient) Then Me.ReleaseBrush
End Sub

'For interop purposes, brushes are often passed around PD as strings.  Any brush can be perfectly re-created from just this string.
Friend Function GetBrushPropertiesAsXML() As String
    
    If (cSerialize Is Nothing) Then Set cSerialize = New pdSerialize
    With cSerialize
        
        'We are now on version 2 of brush parameters (yay?)
        .Reset 2#
        
        Const MAP_255_TO_100 As Single = 1! / 2.55!
        
        .AddParam "brush-mode", Drawing2D.XML_GetNameOfBrushMode(m_BrushMode), True, True
        .AddParam "brush-color", RGB(m_BrushRGBA.Red, m_BrushRGBA.Green, m_BrushRGBA.Blue), True, True
        .AddParam "brush-opacity", CSng(m_BrushRGBA.Alpha) * MAP_255_TO_100, True, True
        .AddParam "brush-pattern-id", Drawing2D.XML_GetNameOfPattern(m_PatternStyle), True, True
        .AddParam "brush-pattern-color-1", RGB(m_PatternRGBA1.Red, m_PatternRGBA1.Green, m_PatternRGBA1.Blue), True, True
        .AddParam "brush-pattern-opacity-1", CSng(m_PatternRGBA1.Alpha) * MAP_255_TO_100, True, True
        .AddParam "brush-pattern-color-2", RGB(m_PatternRGBA2.Red, m_PatternRGBA2.Green, m_PatternRGBA2.Blue), True, True
        .AddParam "brush-pattern-opacity-2", CSng(m_PatternRGBA2.Alpha) * MAP_255_TO_100, True, True
        
        'All other gradient parameters derive from this central string, so we do not need to set them individually
        If (Not m_Gradient Is Nothing) Then
            .AddParam "brush-gradient", m_Gradient.GetGradientAsString(), True, False
        End If
        
        'Textures themselves are not stored.  A solution to this is TBD, but I'm honestly not thrilled about the notion of
        ' serializing an entire texture (which may be enormous) to Base-64.  Texture brushes may just be an exception to
        ' the rule, and you'll be forced to always create them manually.  IDK.
        .AddParam "brush-wrap-mode", Drawing2D.XML_GetNameOfWrapMode(m_TextureWrapMode), True, True
        
    End With
    
    GetBrushPropertiesAsXML = cSerialize.GetParamString()
    
End Function

Friend Sub SetBrushPropertiesFromXML(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If (LenB(srcString) = 0) Then
        Me.ResetAllProperties
    Else
        
        If (cSerialize Is Nothing) Then Set cSerialize = New pdSerialize
        cSerialize.SetParamString srcString
        
        'Check for legacy XML patterns
        If (Not cSerialize.DoesParamExist("brush-mode", True)) Then
            SetBrushPropertiesFromXML_Legacy srcString
            Exit Sub
        End If
        
        'This is a modern XML entry.
        With cSerialize
            m_BrushMode = Drawing2D.XML_GetBrushModeFromName(.GetString("brush-mode", , True))
            Me.SetBrushColor .GetLong("brush-color", vbWhite, True)
            Me.SetBrushOpacity .GetSingle("brush-opacity", 100!, True)
            Me.SetBrushPatternStyle Drawing2D.XML_GetPatternFromName(.GetString("brush-pattern-id", , True))
            Me.SetBrushPattern1Color .GetLong("brush-pattern-color-1", vbWhite, True)
            Me.SetBrushPattern1Opacity .GetSingle("brush-pattern-opacity-1", 100!, True)
            Me.SetBrushPattern2Color .GetLong("brush-pattern-color-2", vbBlack, True)
            Me.SetBrushPattern2Opacity .GetSingle("brush-pattern-opacity-2", 100!, True)
            
            'All other gradient parameters derive from this central string, so we do not need to retrieve them individually
            If .DoesParamExist("brush-gradient") Then
                If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
                m_Gradient.CreateGradientFromString .GetString("brush-gradient", vbNullString)
            End If
            
            'Texture brush *properties* are retrieved, but at present, the texture itself cannot be serialized.
            m_TextureWrapMode = Drawing2D.XML_GetWrapModeFromName(.GetString("brush-wrap-mode", , True))
            
        End With
        
    End If
    
End Sub

Friend Sub SetBrushPropertiesFromXML_Legacy(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If (LenB(srcString) = 0) Then
        Me.ResetAllProperties
    Else
        
        If (cSerialize Is Nothing) Then Set cSerialize = New pdSerialize
        With cSerialize
        
            .SetParamString srcString
            m_BrushMode = .GetLong("BrushMode", P2_BM_Solid)
            Me.SetBrushColor .GetLong("BrushPrimaryColor", vbWhite)
            Me.SetBrushOpacity .GetSingle("BrushPrimaryOpacity", 100!)
            Me.SetBrushPatternStyle .GetLong("BrushPatternID", P2_PS_Horizontal)
            Me.SetBrushPattern1Color .GetLong("BrushPatternColor1", vbWhite)
            Me.SetBrushPattern1Opacity .GetSingle("BrushPatternColor1Opacity", 100!)
            Me.SetBrushPattern2Color .GetLong("BrushPatternColor2", vbBlack)
            Me.SetBrushPattern2Opacity .GetSingle("BrushPatternColor2Opacity", 100!)
            
            'All other gradient parameters derive from this central string, so we do not need to retrieve them individually
            If .DoesParamExist("BrushGradientString") Then
                If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
                m_Gradient.CreateGradientFromString .GetString("BrushGradientString", vbNullString)
            End If
            
            'Texture brush *properties* are retrieved, but at present, the texture itself cannot be serialized.
            m_TextureWrapMode = .GetLong("BrushTextureWrapMode", P2_WM_Tile)
            
        End With
        
    End If
    
End Sub

Friend Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (m_BrushHandle <> 0) Then
        GetHandle = m_BrushHandle
    ElseIf createAsNecessary Then
        If CreateBrush() Then GetHandle = m_BrushHandle
    End If
End Function

Friend Function HasBrush() As Boolean
    HasBrush = (m_BrushHandle <> 0)
End Function

'Create an actual brush handle using the current backend and the current brush settings.
' NOTE: the caller doesn't actually *need* to call this directly.  If GetBrushHandle is called and the brush doesn't yet exist,
'       it will be auto-created.
' NOTE: specialty brushes like the gradient brush may require additional properties to be set.  FOR EXAMPLE, gradient brushes
'       require a boundary rect to know how to scale the gradient - if you haven't set one, this function will fail, or return
'       undesirable results.
Friend Function CreateBrush() As Boolean
    
    If (m_BrushHandle <> 0) Then Me.ReleaseBrush
    
    'VB6 doesn't like passing UDTs ByVal; we use a cheap workaround
    Dim tmpLong As Long, tmpLong2 As Long
    
    Select Case m_BrushMode
    
        'Solid fill
        Case P2_BM_Solid
            GetMem4 VarPtr(m_BrushRGBA), tmpLong
            CreateBrush = (GdipCreateSolidFill(tmpLong, m_BrushHandle) = GP_OK)
            If CreateBrush Then CreateBrush = (m_BrushHandle <> 0)
            
        'Pattern fill
        Case P2_BM_Pattern
            
            'We need to pre-multiply individual pattern opacity values by the brush-wide opacity value before
            ' creating the brush.
            Dim tmpPatternRGBA1 As RGBQuad, tmpPatternRGBA2 As RGBQuad
            tmpPatternRGBA1 = m_PatternRGBA1
            tmpPatternRGBA2 = m_PatternRGBA2
            
            Dim newOpacity As Single
            newOpacity = (CSng(tmpPatternRGBA1.Alpha) / 255!) * (CSng(m_BrushRGBA.Alpha) / 255!)
            tmpPatternRGBA1.Alpha = Int(newOpacity * 255! + 0.5!)
            newOpacity = (CSng(tmpPatternRGBA2.Alpha) / 255!) * (CSng(m_BrushRGBA.Alpha) / 255!)
            tmpPatternRGBA2.Alpha = Int(newOpacity * 255! + 0.5!)
            
            GetMem4 VarPtr(tmpPatternRGBA1), tmpLong
            GetMem4 VarPtr(tmpPatternRGBA2), tmpLong2
            CreateBrush = (GdipCreateHatchBrush(m_PatternStyle, tmpLong, tmpLong2, m_BrushHandle) = GP_OK)
            
        'Gradient fill
        Case P2_BM_Gradient
            If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
            m_BrushHandle = m_Gradient.GetBrushHandle(m_BoundaryRect, False, CSng(m_BrushRGBA.Alpha) / 255!)
            CreateBrush = (m_BrushHandle <> 0)
            
        'Texture fill (*not* serialized - works at run-time only!)
        Case P2_BM_Texture
            
            'Because of the way GDI+ texture brushes work, it is significantly easier to initialize
            ' one from a full DIB object (which *always* guarantees bitmap bits will be available)
            ' vs a GDI+ Graphics object, which is more like a DC in that it could be a non-bitmap
            ' (metafile), dimensionless, etc.
            Dim srcBitmap As Long, tmpReturn As GP_Result
            If GDI_Plus.GetGdipBitmapHandleFromDIB(srcBitmap, m_TextureSrc) Then
                tmpReturn = GdipCreateTexture(srcBitmap, m_TextureWrapMode, m_BrushHandle)
                If (tmpReturn <> GP_OK) Then InternalError "CreateBrush", "GDI+ failure", tmpReturn
                tmpReturn = GdipDisposeImage(srcBitmap)
                If (tmpReturn <> GP_OK) Then InternalError "CreateBrush", "GDI+ failure", tmpReturn
                CreateBrush = (m_BrushHandle <> 0)
            Else
                InternalError "CreateBrush", "GDI+ failure", tmpReturn
            End If
            
    End Select
    
    'When debug mode is active, all object creations are reported back to the central Drawing2D module
    If (CreateBrush And PD2D_DEBUG_MODE) Then Drawing2D.DEBUG_NotifyBrushCountChange True
    
End Function

Friend Function ReleaseBrush() As Boolean
    
    If (m_BrushHandle <> 0) Then
        
        'Call the backend-specific release function
        ReleaseBrush = (GdipDeleteBrush(m_BrushHandle) = GP_OK)
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseBrush Then
            m_BrushHandle = 0
            If PD2D_DEBUG_MODE Then Drawing2D.DEBUG_NotifyBrushCountChange False
        End If
    
    Else
        ReleaseBrush = True
    End If
    
End Function

Friend Sub ResetAllProperties()
    
    Me.SetBrushMode P2_BM_Solid
    Me.SetBrushColor vbWhite
    Me.SetBrushOpacity 100!
    Me.SetBrushPatternStyle P2_PS_Horizontal
    Me.SetBrushPattern1Color vbWhite
    Me.SetBrushPattern1Opacity 100!
    Me.SetBrushPattern2Color vbBlack
    Me.SetBrushPattern2Opacity 100!
    
    'All other gradient parameters derive from this central string, so we do not need to set them individually
    Me.SetBrushGradientAllSettings vbNullString
    
    Me.SetBrushTextureWrapMode P2_WM_Tile
    
    'Free any associated textures
    Set m_TextureSrc = Nothing
    
End Sub

Private Sub Class_Initialize()
    
    Me.ResetAllProperties
    
    'Prep a default boundary rect
    With m_BoundaryRect
        .Left = 0!
        .Top = 0!
        .Width = 100!
        .Height = 100!
    End With
    
End Sub

Private Sub Class_Terminate()
    Me.ReleaseBrush
End Sub

'All pd2D classes report errors using an internal function similar to this one.
' Feel free to modify this function to better fit your project
' (for example, maybe you prefer to raise an actual error event).
'
'Note that by default, pd2D build simply dumps all error information to the Immediate window.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    Drawing2D.DEBUG_NotifyError "pd2DBrush", errFunction, errDescription, errNum
End Sub
