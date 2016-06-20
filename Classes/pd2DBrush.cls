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
'Copyright 2014-2016 by Tanner Helland
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
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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
Private m_BoundaryRect As RECTF

'Texture brushes are much simpler if we maintain a copy of the source texture (e.g. this avoids messy cases like the caller
' creating a texture brush, then erasing the source image prior to freeing *this* brush).  For simplicity, the brush is stored
' inside a pdDIB object, with any clipping applied *prior* to creating the DIB.
Private m_TextureSrc As pdDIB
Private m_TextureWrapMode As PD_2D_WrapMode

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML

'The backend used to render this brush.  At present, this is immutable; once set, it cannot be migrated.
Private m_BrushBackend As PD_2D_RENDERING_BACKEND

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Once a brush has been created, this handle value will be non-zero
Private m_BrushHandle As Long

'NOTE: texture brush settings are still TBD

'Get/set individual settings.  Note that these just wrap the generic Get/SetBrushProperty functions, below.
Public Function GetBrushMode() As PD_2D_BrushMode
    GetBrushMode = Me.GetBrushProperty(P2_BrushMode)
End Function

Public Function GetBrushColor() As Long
    GetBrushColor = Me.GetBrushProperty(P2_BrushColor)
End Function

Public Function GetBrushOpacity() As Single
    GetBrushOpacity = Me.GetBrushProperty(P2_BrushOpacity)
End Function

Public Function GetBrushPatternStyle() As PD_2D_PatternStyle
    GetBrushPatternStyle = Me.GetBrushProperty(P2_BrushPatternStyle)
End Function

Public Function GetBrushPattern1Color() As Long
    GetBrushPattern1Color = Me.GetBrushProperty(P2_BrushPattern1Color)
End Function

Public Function GetBrushPattern1Opacity() As Single
    GetBrushPattern1Opacity = Me.GetBrushProperty(P2_BrushPattern1Opacity)
End Function

Public Function GetBrushPattern2Color() As Long
    GetBrushPattern2Color = Me.GetBrushProperty(P2_BrushPattern2Color)
End Function

Public Function GetBrushPattern2Opacity() As Single
    GetBrushPattern2Opacity = Me.GetBrushProperty(P2_BrushPattern2Opacity)
End Function

Public Function GetBrushGradientAllSettings() As String
    GetBrushGradientAllSettings = Me.GetBrushProperty(P2_BrushGradientAllSettings)
End Function
    
Friend Function GetBrushGradientShape() As PD_2D_GradientShape
    GetBrushGradientShape = Me.GetBrushProperty(P2_BrushGradientShape)
End Function

Public Function GetBrushGradientAngle() As Single
    GetBrushGradientAngle = Me.GetBrushProperty(P2_BrushGradientAngle)
End Function

Public Function GetBrushGradientWrapMode() As PD_2D_WrapMode
    GetBrushGradientWrapMode = Me.GetBrushProperty(P2_BrushGradientAngle)
End Function

Public Function GetBrushGradientNodes() As String
    GetBrushGradientNodes = Me.GetBrushProperty(P2_BrushGradientAngle)
End Function

Public Function GetBrushTextureWrapMode() As PD_2D_WrapMode
    GetBrushTextureWrapMode = Me.GetBrushProperty(P2_BrushTextureWrapMode)
End Function

Public Sub SetBrushMode(ByVal newSetting As PD_2D_BrushMode)
    Me.SetBrushProperty P2_BrushMode, newSetting
End Sub

Public Sub SetBrushColor(ByVal newSetting As Long)
    Me.SetBrushProperty P2_BrushColor, newSetting
End Sub

Public Sub SetBrushOpacity(ByVal newSetting As Single)
    Me.SetBrushProperty P2_BrushOpacity, newSetting
End Sub

Public Sub SetBrushPatternStyle(ByVal newSetting As PD_2D_PatternStyle)
    Me.SetBrushProperty P2_BrushPatternStyle, newSetting
End Sub

Public Sub SetBrushPattern1Color(ByVal newSetting As Long)
    Me.SetBrushProperty P2_BrushPattern1Color, newSetting
End Sub

Public Sub SetBrushPattern1Opacity(ByVal newSetting As Single)
    Me.SetBrushProperty P2_BrushPattern1Opacity, newSetting
End Sub

Public Sub SetBrushPattern2Color(ByVal newSetting As Long)
    Me.SetBrushProperty P2_BrushPattern2Color, newSetting
End Sub

Public Sub SetBrushPattern2Opacity(ByVal newSetting As Single)
    Me.SetBrushProperty P2_BrushPattern2Opacity, newSetting
End Sub

Public Sub SetBrushGradientAllSettings(ByVal newSetting As String)
    Me.SetBrushProperty P2_BrushGradientAllSettings, newSetting
End Sub
    
Friend Sub SetBrushGradientShape(ByVal newSetting As PD_2D_GradientShape)
    Me.SetBrushProperty P2_BrushGradientShape, newSetting
End Sub

Public Sub SetBrushGradientAngle(ByVal newSetting As Single)
    Me.SetBrushProperty P2_BrushGradientAngle, newSetting
End Sub

Public Sub SetBrushGradientWrapMode(ByVal newSetting As PD_2D_WrapMode)
    Me.SetBrushProperty P2_BrushGradientAngle, newSetting
End Sub

Public Sub SetBrushGradientNodes(ByVal newSetting As String)
    Me.SetBrushProperty P2_BrushGradientAngle, newSetting
End Sub

Public Sub SetBrushTextureFromDC(ByVal srcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcColorDepth As Long = 32)
    If (m_TextureSrc Is Nothing) Then Set m_TextureSrc = New pdDIB
    m_TextureSrc.CreateFromDC srcDC, srcX, srcY, srcWidth, srcHeight, srcColorDepth
End Sub

Public Sub SetBrushTextureFromDIB(ByVal srcDIB As pdDIB)
    If (m_TextureSrc Is Nothing) Then Set m_TextureSrc = New pdDIB
    m_TextureSrc.CreateFromExistingDIB srcDIB
End Sub

Public Sub SetBrushTextureWrapMode(ByVal newSetting As PD_2D_WrapMode)
    Me.SetBrushProperty P2_BrushTextureWrapMode, newSetting
End Sub

'Get/set individual settings by enum type
Public Function GetBrushProperty(ByVal propID As PD_2D_BRUSH_SETTINGS) As Variant
    
    Select Case propID
        
        Case P2_BrushMode
            GetBrushProperty = m_BrushMode
        
        Case P2_BrushColor
            GetBrushProperty = m_PrimaryColor
            
        Case P2_BrushOpacity
            GetBrushProperty = m_PrimaryOpacity
            
        Case P2_BrushPatternStyle
            GetBrushProperty = m_PatternStyle
            
        Case P2_BrushPattern1Color
            GetBrushProperty = m_Pattern1Color
            
        Case P2_BrushPattern1Opacity
            GetBrushProperty = m_Pattern1Opacity
            
        Case P2_BrushPattern2Color
            GetBrushProperty = m_Pattern2Color
            
        Case P2_BrushPattern2Opacity
            GetBrushProperty = m_Pattern2Opacity
            
        Case P2_BrushGradientAllSettings
            GetBrushProperty = m_Gradient.GetGradientAsString()
            
        Case P2_BrushGradientShape
            GetBrushProperty = m_Gradient.GetGradientShape()
            
        Case P2_BrushGradientAngle
            GetBrushProperty = m_Gradient.GetGradientAngle()
            
        Case P2_BrushGradientWrapMode
            GetBrushProperty = m_Gradient.GetGradientWrapMode()
            
        Case P2_BrushGradientNodes
            GetBrushProperty = m_Gradient.GetGradientNodes()
            
        Case P2_BrushTextureWrapMode
            GetBrushProperty = m_TextureWrapMode
            
    End Select
    
End Function

Public Sub SetBrushProperty(ByVal propID As PD_2D_BRUSH_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
        
        Case P2_BrushMode
            m_BrushMode = CLng(newSetting)
        
        Case P2_BrushColor
            m_PrimaryColor = CLng(newSetting)
            
        Case P2_BrushOpacity
            m_PrimaryOpacity = CDbl(newSetting)
            
        Case P2_BrushPatternStyle
            m_PatternStyle = CLng(newSetting)
            
        Case P2_BrushPattern1Color
            m_Pattern1Color = CLng(newSetting)
            
        Case P2_BrushPattern1Opacity
            m_Pattern1Opacity = CDbl(newSetting)
            
        Case P2_BrushPattern2Color
            m_Pattern2Color = CLng(newSetting)
            
        Case P2_BrushPattern2Opacity
            m_Pattern2Opacity = CDbl(newSetting)
            
        Case P2_BrushGradientAllSettings
            If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
            m_Gradient.CreateGradientFromString newSetting
        
        Case P2_BrushGradientShape
            If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
            m_Gradient.SetGradientShape CLng(newSetting)
            
        Case P2_BrushGradientAngle
            If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
            m_Gradient.SetGradientAngle CSng(newSetting)
            
        Case P2_BrushGradientWrapMode
            If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
            m_Gradient.SetGradientWrapMode CLng(newSetting)
            
        Case P2_BrushGradientNodes
            If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
            m_Gradient.SetGradientNodes CStr(newSetting)
            
        Case P2_BrushTextureWrapMode
            m_TextureWrapMode = CLng(newSetting)
            
    End Select
    
    RelayPropertyChangeToBackend propID
    
End Sub

'NOTE: some brush setting changes require us to completely recreate the brush.  (GDI+ does not support changing certain
'      settings post-creation.)  This function will handle brush re-creation automatically, but be advised that this will
'      screw with attempts to manually cache brushes.  Instead, always rely on the GetHandle() function, as it may return
'      different handles between property changes.
Private Sub RelayPropertyChangeToBackend(ByVal propID As PD_2D_BRUSH_SETTINGS)
    
    'Property changes only need to be relayed if we've already created a handle
    If (m_BrushHandle <> 0) Then
    
        'Relays obviously vary by the active backend!
        Select Case m_BrushBackend
            
            Case P2_DefaultBackend, P2_GDIPlusBackend
                
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
                        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientAllSettings, m_Gradient.GetGradientAsString())) Then ReleaseBrush
                    
                    Case P2_BrushGradientShape
                        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientShape, m_Gradient.GetGradientShape())) Then ReleaseBrush
                        
                    Case P2_BrushGradientAngle
                        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientAngle, m_Gradient.GetGradientAngle())) Then ReleaseBrush
                        
                    Case P2_BrushGradientWrapMode
                        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientWrapMode, m_Gradient.GetGradientWrapMode())) Then ReleaseBrush
                        
                    Case P2_BrushGradientNodes
                        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushGradientNodes, m_Gradient.GetGradientNodes())) Then ReleaseBrush
                    
                    Case P2_BrushTextureWrapMode
                        If (Not GDI_Plus.SetGDIPlusBrushProperty(m_BrushHandle, P2_BrushTextureWrapMode, m_TextureWrapMode)) Then ReleaseBrush
                    
                End Select
                
            'In the future, other backends can be handled here...
            Case Else
                
        End Select
    
    End If
        
End Sub

Friend Sub SetBoundaryRect(ByRef srcRect As RECTF)
    m_BoundaryRect = srcRect
    If (m_BrushMode = P2_BM_Gradient) Then ReleaseBrush
End Sub

'For interop purposes, brushes are often passed around PD as strings.  Any brush can be perfectly re-created from just this string.
Public Function GetBrushPropertiesAsXML() As String
    
    If (cSerialize Is Nothing) Then Set cSerialize = New pdParamXML
    With cSerialize
        .Reset 1#
        .AddParam "BrushMode", m_BrushMode
        .AddParam "BrushPrimaryColor", m_PrimaryColor
        .AddParam "BrushPrimaryOpacity", m_PrimaryOpacity
        .AddParam "BrushPatternID", m_PatternStyle
        .AddParam "BrushPatternColor1", m_Pattern1Color
        .AddParam "BrushPatternColor1Opacity", m_Pattern1Opacity
        .AddParam "BrushPatternColor2", m_Pattern2Color
        .AddParam "BrushPatternColor2Opacity", m_Pattern2Opacity
        
        'All other gradient parameters derive from this master string, so we do not need to set them individually
        .AddParam "BrushGradientString", m_Gradient.GetGradientAsString()
        
        'Textures themselves are not stored.  A solution to this is TBD, but I'm honestly not thrilled about the notion of
        ' serializing an entire texture (which may be enormous) to Base-64.  Texture brushes may just be an exception to
        ' the rule, and you'll be forced to always create them manually.  IDK.
        .AddParam "BrushTextureWrapMode", m_TextureWrapMode
    End With
    
    GetBrushPropertiesAsXML = cSerialize.GetParamString
    
End Function

Public Sub SetBrushPropertiesFromXML(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If Len(srcString) = 0 Then
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
            m_Gradient.CreateGradientFromString .GetString("BrushGradientString", vbNullString)
            
            'Texture brush *properties* are retrieved, but at present, the texture itself cannot be serialized.
            m_TextureWrapMode = .GetLong("BrushTextureWrapMode", P2_WM_Tile)
        End With
        
    End If
    
End Sub

Public Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_BrushHandle = 0)) Then
        If CreateBrush() Then GetHandle = m_BrushHandle Else GetHandle = 0
    Else
        GetHandle = m_BrushHandle
    End If
End Function

Public Function GetBrushBackend() As PD_2D_RENDERING_BACKEND
    GetBrushBackend = m_BrushBackend
End Function

Public Function HasBrush() As Boolean
    HasBrush = CBool(m_BrushHandle <> 0)
End Function

'You only need to manually initialize a brush backend if you plan on using a non-default backend.  (Brushes with the
' default backend are initialized automatically.)
Friend Function InitializeBrushBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_BrushBackend = targetBackend
    InitializeBrushBackend = True
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
            tmpOpacityModifier = m_PrimaryOpacity / 100
            m_BrushHandle = GDI_Plus.GetGDIPlusPatternBrushHandle(m_PatternStyle, m_Pattern1Color, m_Pattern1Opacity * 2.55 * tmpOpacityModifier, m_Pattern2Color, m_Pattern2Opacity * 2.55 * tmpOpacityModifier)
            
        'Gradient fill
        Case P2_BM_Gradient
            m_BrushHandle = m_Gradient.GetBrushHandle(m_BoundaryRect, False, m_PrimaryOpacity / 100, m_BrushBackend)
            
        'Texture fill (TODO)
        Case P2_BM_Texture
            m_BrushHandle = GDI_Plus.GetGDIPlusTextureBrush(m_TextureSrc, m_TextureWrapMode)
            
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    CreateBrush = CBool(m_BrushHandle <> 0)
    If (CreateBrush And m_DebugMode) Then Drawing2D.DEBUG_NotifyBrushCountChange m_BrushBackend, True
    
End Function

Public Function ReleaseBrush() As Boolean
    
    If (m_BrushHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_BrushBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleaseBrush = GDI_Plus.ReleaseGDIPlusBrush(m_BrushHandle)
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseBrush Then
            m_BrushHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyBrushCountChange m_BrushBackend, False
        End If
    
    Else
        ReleaseBrush = True
    End If
    
End Function

Public Sub ResetAllProperties()

    Me.SetBrushProperty P2_BrushMode, P2_BM_Solid
    Me.SetBrushProperty P2_BrushColor, vbWhite
    Me.SetBrushProperty P2_BrushOpacity, 100#
    Me.SetBrushProperty P2_BrushPatternStyle, P2_PS_Horizontal
    Me.SetBrushProperty P2_BrushPattern1Color, vbWhite
    Me.SetBrushProperty P2_BrushPattern1Opacity, 100#
    Me.SetBrushProperty P2_BrushPattern2Color, vbBlack
    Me.SetBrushProperty P2_BrushPattern2Opacity, 100#
    
    'All other gradient parameters derive from this master string, so we do not need to set them individually
    Me.SetBrushProperty P2_BrushGradientAllSettings, vbNullString
    
    Me.SetBrushProperty P2_BrushTextureWrapMode, P2_WM_Tile
    
    'Free any associated textures
    Set m_TextureSrc = Nothing
    
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    
    Me.ResetAllProperties
    
    'Prep a default gradient handler and gradient-related properties
    Set m_Gradient = New pd2DGradient
    With m_BoundaryRect
        .Left = 0
        .Top = 0
        .Width = 100
        .Height = 100
    End With
    
End Sub

Private Sub Class_Terminate()
    Me.ReleaseBrush
End Sub
