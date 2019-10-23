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
'Copyright 2014-2019 by Tanner Helland
'Created: 30/June/15 (but assembled from many bits written earlier)
'Last updated: 13/May/16
'Last update: total overhaul
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit


'Current brush mode (solid fill, gradient, texture, etc)
Private m_BrushMode As PD_2D_BrushMode

'Solid brush settings
Private m_PrimaryColor As Long
Private m_PrimaryOpacity As Single

'Pattern brush settings
Private m_PatternStyle As PD_2D_PatternStyle
Private m_Pattern1Color As Long
Private m_Pattern1Opacity As Single
Private m_Pattern2Color As Long
Private m_Pattern2Opacity As Single

'Gradient brushes are more complicated, as they rely on a variable number of settings.
' We use a specialized class to manage those settings.
Private m_Gradient As pd2DGradient

'Gradient fills also require a boundary rect to define the start/end positions of the gradient line.  You *must set this prior
' to requesting a gradient brush* or the gradient results will look funny!
Private m_BoundaryRect As RectF

'Texture brushes are much simpler if we maintain a copy of the source texture (e.g. this avoids messy cases like the caller
' creating a texture brush, then erasing the source image prior to freeing *this* brush).  For simplicity, the brush is stored
' inside a pdDIB object, with any clipping applied *prior* to creating the DIB.
Private m_TextureSrc As pdDIB
Private m_TextureWrapMode As PD_2D_WrapMode

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Once a brush has been created, this handle value will be non-zero
Private m_BrushHandle As Long

'NOTE: texture brush settings are still TBD

'Get/set individual settings.
Friend Function GetBrushMode() As PD_2D_BrushMode
    GetBrushMode = m_BrushMode
End Function

Friend Function GetBrushColor() As Long
    GetBrushColor = m_PrimaryColor
End Function

Friend Function GetBrushOpacity() As Single
    GetBrushOpacity = m_PrimaryOpacity
End Function

Friend Function GetBrushPatternStyle() As PD_2D_PatternStyle
    GetBrushPatternStyle = m_PatternStyle
End Function

Friend Function GetBrushPattern1Color() As Long
    GetBrushPattern1Color = m_Pattern1Color
End Function

Friend Function GetBrushPattern1Opacity() As Single
    GetBrushPattern1Opacity = m_Pattern1Opacity
End Function

Friend Function GetBrushPattern2Color() As Long
    GetBrushPattern2Color = m_Pattern2Color
End Function

Friend Function GetBrushPattern2Opacity() As Single
    GetBrushPattern2Opacity = m_Pattern2Opacity
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
    
    If (m_BrushMode <> newSetting) Then
        ReleaseBrush
        m_BrushMode = newSetting
    End If
    
End Sub

Friend Sub SetBrushColor(ByVal newSetting As Long)
    
    m_PrimaryColor = newSetting
    
    'After setting a parameter, if a brush already exists, we must relay the new setting to the active backend
    If (m_BrushHandle <> 0) Then
        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushColor, m_PrimaryColor)) Then ReleaseBrush
    End If
    
End Sub

Friend Sub SetBrushOpacity(ByVal newSetting As Single)
    
    m_PrimaryOpacity = newSetting
    
    If (m_BrushHandle <> 0) Then
        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushOpacity, m_PrimaryOpacity)) Then ReleaseBrush
    End If
    
End Sub

Friend Sub SetBrushPatternStyle(ByVal newSetting As PD_2D_PatternStyle)
    
    m_PatternStyle = newSetting
    
    'We don't need to relay certain pattern settings to GDI+, as this class handles that data internally
    
End Sub

Friend Sub SetBrushPattern1Color(ByVal newSetting As Long)
    
    m_Pattern1Color = newSetting
    
    'We don't need to relay certain pattern settings to GDI+, as this class handles that data internally
    
End Sub

Friend Sub SetBrushPattern1Opacity(ByVal newSetting As Single)
    
    m_Pattern1Opacity = newSetting
    
    'We don't need to relay certain pattern settings to GDI+, as this class handles that data internally
    
End Sub

Friend Sub SetBrushPattern2Color(ByVal newSetting As Long)
    
    m_Pattern2Color = newSetting
    
    'We don't need to relay certain pattern settings to GDI+, as this class handles that data internally
    
End Sub

Friend Sub SetBrushPattern2Opacity(ByVal newSetting As Single)
    
    m_Pattern2Opacity = newSetting
    
    'We don't need to relay certain pattern settings to GDI+, as this class handles that data internally
    
End Sub

Friend Sub SetBrushGradientAllSettings(ByRef newSetting As String)

    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.CreateGradientFromString newSetting
    
    'We don't need to relay gradient settings to GDI+, as our gradient class handles that data internally
    
End Sub
    
Friend Sub SetBrushGradientShape(ByVal newSetting As PD_2D_GradientShape)
    
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientShape newSetting
    
    'We don't need to relay gradient settings to GDI+, as our gradient class handles that data internally
    
End Sub

Friend Sub SetBrushGradientAngle(ByVal newSetting As Single)
    
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientAngle newSetting
    
    'We don't need to relay gradient settings to GDI+, as our gradient class handles that data internally
    
End Sub

Friend Sub SetBrushGradientWrapMode(ByVal newSetting As PD_2D_WrapMode)
    
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientWrapMode newSetting
    
    'We don't need to relay gradient settings to GDI+, as our gradient class handles that data internally
    
End Sub

Friend Sub SetBrushGradientNodes(ByRef newSetting As String)
    
    If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
    m_Gradient.SetGradientNodes newSetting
    
    'We don't need to relay gradient settings to GDI+, as our gradient class handles that data internally
    
End Sub

Friend Sub SetBrushTextureFromDC(ByVal srcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcColorDepth As Long = 32)
    If (m_TextureSrc Is Nothing) Then Set m_TextureSrc = New pdDIB
    m_TextureSrc.CreateFromDC srcDC, srcX, srcY, srcWidth, srcHeight, srcColorDepth
End Sub

Friend Sub SetBrushTextureFromDIB(ByRef srcDIB As pdDIB)
    If (m_TextureSrc Is Nothing) Then Set m_TextureSrc = New pdDIB
    m_TextureSrc.CreateFromExistingDIB srcDIB
End Sub

'Unlike other brush functions, this one requires that the brush has already been created.
' (This is due to the way the transform handle is cached, and our inability to ensure the lifetime
' of the source transform object.)
Friend Sub SetBrushTextureTransform(ByRef srcTransform As pd2DTransform)
    If (m_BrushHandle <> 0) Then
        If (Not GDI_Plus.GDIPlus_SetTextureBrushTransform(m_BrushHandle, srcTransform.GetHandle)) Then InternalError "SetBrushTextureTransform failed"
    Else
        InternalError "Brush must exist before calling SetBrushTextureTransform"
    End If
End Sub

Friend Sub SetBrushTextureWrapMode(ByVal newSetting As PD_2D_WrapMode)
    
    m_TextureWrapMode = newSetting
    
    If (m_BrushHandle <> 0) Then
        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushTextureWrapMode, m_TextureWrapMode)) Then ReleaseBrush
    End If
    
End Sub

'NOTE: some brush setting changes require us to completely recreate the brush.  (GDI+ does not support changing certain
'      settings post-creation.)  This function will handle brush re-creation automatically, but be advised that this will
'      screw with attempts to manually cache brushes.  Instead, always rely on the GetHandle() function, as it may return
'      different handles between property changes.
Private Sub RelayPropertyChangeToBackend(ByVal propID As PD_2D_BRUSH_SETTINGS)
    
    'Property changes only need to be relayed if we've already created a handle
    If (m_BrushHandle <> 0) Then
    
        Select Case propID
                        
            Case P2_BrushMode
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushMode, m_BrushMode)) Then ReleaseBrush
            
            Case P2_BrushColor
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushColor, m_PrimaryColor)) Then ReleaseBrush
                
            Case P2_BrushOpacity
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushOpacity, m_PrimaryOpacity)) Then ReleaseBrush
                
            Case P2_BrushPatternStyle
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushPatternStyle, m_PatternStyle)) Then ReleaseBrush
                
            Case P2_BrushPattern1Color
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushPattern1Color, m_Pattern1Color)) Then ReleaseBrush
                
            Case P2_BrushPattern1Opacity
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushPattern1Opacity, m_Pattern1Opacity)) Then ReleaseBrush
                
            Case P2_BrushPattern2Color
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushPattern2Color, m_Pattern2Color)) Then ReleaseBrush
                
            Case P2_BrushPattern2Opacity
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushPattern2Opacity, m_Pattern2Opacity)) Then ReleaseBrush
                
            Case P2_BrushGradientAllSettings
                If (Not m_Gradient Is Nothing) Then
                    If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientAllSettings, m_Gradient.GetGradientAsString())) Then ReleaseBrush
                End If
            
            Case P2_BrushGradientShape
                If (Not m_Gradient Is Nothing) Then
                    If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientShape, m_Gradient.GetGradientShape())) Then ReleaseBrush
                End If
                
            Case P2_BrushGradientAngle
                If (Not m_Gradient Is Nothing) Then
                    If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientAngle, m_Gradient.GetGradientAngle())) Then ReleaseBrush
                End If
                
            Case P2_BrushGradientWrapMode
                If (Not m_Gradient Is Nothing) Then
                    If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientWrapMode, m_Gradient.GetGradientWrapMode())) Then ReleaseBrush
                End If
                
            Case P2_BrushGradientNodes
                If (Not m_Gradient Is Nothing) Then
                    If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientNodes, m_Gradient.GetGradientNodes())) Then ReleaseBrush
                End If
            
            Case P2_BrushTextureWrapMode
                If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushTextureWrapMode, m_TextureWrapMode)) Then ReleaseBrush
            
        End Select
        
    End If
        
End Sub

Friend Sub SetBoundaryRect(ByRef srcRect As RectF)
    m_BoundaryRect = srcRect
    If (m_BrushMode = P2_BM_Gradient) Then ReleaseBrush
End Sub

'For interop purposes, brushes are often passed around PD as strings.  Any brush can be perfectly re-created from just this string.
Friend Function GetBrushPropertiesAsXML() As String
    
    If (cSerialize Is Nothing) Then Set cSerialize = New pdParamXML
    With cSerialize
        .Reset 1#
        .AddParam "BrushMode", m_BrushMode, True
        .AddParam "BrushPrimaryColor", m_PrimaryColor, True
        .AddParam "BrushPrimaryOpacity", m_PrimaryOpacity, True
        .AddParam "BrushPatternID", m_PatternStyle, True
        .AddParam "BrushPatternColor1", m_Pattern1Color, True
        .AddParam "BrushPatternColor1Opacity", m_Pattern1Opacity, True
        .AddParam "BrushPatternColor2", m_Pattern2Color, True
        .AddParam "BrushPatternColor2Opacity", m_Pattern2Opacity, True
        
        'All other gradient parameters derive from this master string, so we do not need to set them individually
        If (Not m_Gradient Is Nothing) Then
            .AddParam "BrushGradientString", m_Gradient.GetGradientAsString(), True
        End If
        
        'Textures themselves are not stored.  A solution to this is TBD, but I'm honestly not thrilled about the notion of
        ' serializing an entire texture (which may be enormous) to Base-64.  Texture brushes may just be an exception to
        ' the rule, and you'll be forced to always create them manually.  IDK.
        .AddParam "BrushTextureWrapMode", m_TextureWrapMode, True
    End With
    
    GetBrushPropertiesAsXML = cSerialize.GetParamString
    
End Function

Friend Sub SetBrushPropertiesFromXML(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If (LenB(srcString) = 0) Then
        Me.ResetAllProperties
    Else
        
        If (cSerialize Is Nothing) Then Set cSerialize = New pdParamXML
        With cSerialize
        
            .SetParamString srcString
            m_BrushMode = .GetLong("BrushMode", P2_BM_Solid)
            m_PrimaryColor = .GetLong("BrushPrimaryColor", vbWhite)
            m_PrimaryOpacity = .GetDouble("BrushPrimaryOpacity", 100#)
            m_PatternStyle = .GetLong("BrushPatternID", P2_PS_Horizontal)
            m_Pattern1Color = .GetLong("BrushPatternColor1", vbWhite)
            m_Pattern1Opacity = .GetDouble("BrushPatternColor1Opacity", 100#)
            m_Pattern2Color = .GetLong("BrushPatternColor2", vbBlack)
            m_Pattern2Opacity = .GetDouble("BrushPatternColor2Opacity", 100#)
            
            'All other gradient parameters derive from this master string, so we do not need to retrieve them individually
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
    If (createAsNecessary And (m_BrushHandle = 0)) Then
        If Me.CreateBrush() Then GetHandle = m_BrushHandle Else GetHandle = 0
    Else
        GetHandle = m_BrushHandle
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
    
    If (m_BrushHandle <> 0) Then ReleaseBrush
    
    Select Case m_BrushMode
    
        'Solid fill
        Case P2_BM_Solid
            m_BrushHandle = GDI_Plus.GetGDIPlusSolidBrushHandle(m_PrimaryColor, m_PrimaryOpacity * 2.55)
            
        'Pattern fill
        Case P2_BM_Pattern
            Dim tmpOpacityModifier As Double
            tmpOpacityModifier = m_PrimaryOpacity * 0.01
            m_BrushHandle = GDI_Plus.GetGDIPlusPatternBrushHandle(m_PatternStyle, m_Pattern1Color, m_Pattern1Opacity * 2.55 * tmpOpacityModifier, m_Pattern2Color, m_Pattern2Opacity * 2.55 * tmpOpacityModifier)
            
        'Gradient fill
        Case P2_BM_Gradient
            If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
            m_BrushHandle = m_Gradient.GetBrushHandle(m_BoundaryRect, False, m_PrimaryOpacity * 0.01)
            
        'Texture fill (TODO)
        Case P2_BM_Texture
            m_BrushHandle = GDI_Plus.GetGDIPlusTextureBrush(m_TextureSrc, m_TextureWrapMode)
            
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    CreateBrush = (m_BrushHandle <> 0)
    If (CreateBrush And m_DebugMode) Then Drawing2D.DEBUG_NotifyBrushCountChange True
    
End Function

Friend Function ReleaseBrush() As Boolean
    
    If (m_BrushHandle <> 0) Then
        
        'Call the backend-specific release function
        ReleaseBrush = GDI_Plus.ReleaseGDIPlusBrush(m_BrushHandle)
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseBrush Then
            m_BrushHandle = 0
            If m_DebugMode Then Drawing2D.DEBUG_NotifyBrushCountChange False
        End If
    
    Else
        ReleaseBrush = True
    End If
    
End Function

Friend Sub ResetAllProperties()
    
    Me.SetBrushMode P2_BM_Solid
    Me.SetBrushColor vbWhite
    Me.SetBrushOpacity 100#
    Me.SetBrushPatternStyle P2_PS_Horizontal
    Me.SetBrushPattern1Color vbWhite
    Me.SetBrushPattern1Opacity 100#
    Me.SetBrushPattern2Color vbBlack
    Me.SetBrushPattern2Opacity 100#
    
    'All other gradient parameters derive from this master string, so we do not need to set them individually
    Me.SetBrushGradientAllSettings vbNullString
    
    Me.SetBrushTextureWrapMode P2_WM_Tile
    
    'Free any associated textures
    Set m_TextureSrc = Nothing
    
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Friend Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
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
    
    m_DebugMode = Drawing2D.GetLibraryDebugMode()
    
End Sub

Private Sub Class_Terminate()
    Me.ReleaseBrush
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DBrush"
End Sub
