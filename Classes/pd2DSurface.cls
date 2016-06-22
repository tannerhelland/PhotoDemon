VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 2D Rendering Surface class
'Copyright 2012-2016 by Tanner Helland
'Created: 01/September/12
'Last updated: 10/May/16
'Last update: split surface-specific capabilities out from GDI+ and into a safe wrapper
'
'In the pd2D drawing model, "Surfaces" are objects onto which you can paint.  At present, there are two types of surfaces:
'
' 1) Wrappers around existing DCs.  This is help for UI painting, as wrapping a DC is effectively instantaneous
'    (compared to making a copy of the DC's bitmap, painting to that, then flipping back to the DC when finished)
'
' 2) Bitmap (raster) surfaces.  The creation and destruction of such surfaces is handled internally, using a pdDIB object.
'
'Regardless of which surface type you use, you can always access an hDC for external GDI interop purposes.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'The current surface type.  At present, this is immutable; once set, it cannot be migrated
Private m_SurfaceBackend As PD_2D_RENDERING_BACKEND

'The current surface handle.  Will be 0 if the surface is uninitialized.
Private m_SurfaceHandle As Long

'If this surface was created by wrapping an existing hDC, this value will be non-zero.
Private m_SurfaceDC As Long

'If you need to query a wrapped DC for bitmap information (e.g. width/height), use this temporary object to do so.
Private m_SurfaceDCBitmap As GDI_Bitmap

'If this surface was created as a new surface, this DIB will serve as the surface backer.  (Similarly, the surface DC will simply
' wrap this DIB's hDC.)
Private m_SurfaceDIB As pdDIB

'I don't currently expose a surface's type; this is managed internally, based on the way the surface is created.
Private m_SurfaceType As PD_2D_SurfaceType

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Surface settings that support get/set operations
Private m_SurfaceAntialiasing As PD_2D_Antialiasing
Private m_SurfacePixelOffset As PD_2D_PixelOffset
Private m_SurfaceRenderingOriginX As Single, m_SurfaceRenderingOriginY As Single
Private m_SurfaceBlendUsingSRGBGamma As Boolean

'Get/set individual settings.  Note that these just wrap the generic Get/SetPenProperty functions, below.
Public Function GetSurfaceAntialiasing() As PD_2D_Antialiasing
    GetSurfaceAntialiasing = Me.GetSurfaceProperty(P2_SurfaceAntialiasing)
End Function

Public Function GetSurfacePixelOffset() As PD_2D_PixelOffset
    GetSurfacePixelOffset = Me.GetSurfaceProperty(P2_SurfacePixelOffset)
End Function

Public Function GetSurfaceRenderingOriginX() As Single
    GetSurfaceRenderingOriginX = Me.GetSurfaceProperty(P2_SurfaceRenderingOriginX)
End Function

Public Function GetSurfaceRenderingOriginY() As Single
    GetSurfaceRenderingOriginY = Me.GetSurfaceProperty(P2_SurfaceRenderingOriginY)
End Function

Public Function GetSurfaceBlendUsingSRGBGamma() As Boolean
    GetSurfaceBlendUsingSRGBGamma = Me.GetSurfaceProperty(P2_SurfaceBlendUsingSRGBGamma)
End Function

Public Sub SetSurfaceAntialiasing(ByVal newSetting As PD_2D_Antialiasing)
    Me.SetSurfaceProperty P2_SurfaceAntialiasing, newSetting
End Sub

Public Sub SetSurfacePixelOffset(ByVal newSetting As PD_2D_PixelOffset)
    Me.SetSurfaceProperty P2_SurfacePixelOffset, newSetting
End Sub

Public Sub SetSurfaceRenderingOriginX(ByVal newSetting As Single)
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginX, newSetting
End Sub

Public Sub SetSurfaceRenderingOriginY(ByVal newSetting As Single)
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginY, newSetting
End Sub

Public Sub SetSurfaceBlendUsingSRGBGamma(ByVal newSetting As Boolean)
    Me.SetSurfaceProperty P2_SurfaceBlendUsingSRGBGamma, newSetting
End Sub

'Get/set individual settings by enum type
Public Function GetSurfaceProperty(ByVal propID As PD_2D_SURFACE_SETTINGS) As Variant

    Select Case propID
    
        Case P2_SurfaceAntialiasing
            GetSurfaceProperty = m_SurfaceAntialiasing
        
        Case P2_SurfacePixelOffset
            GetSurfaceProperty = m_SurfacePixelOffset
            
        Case P2_SurfaceRenderingOriginX
            GetSurfaceProperty = m_SurfaceRenderingOriginX
        
        Case P2_SurfaceRenderingOriginY
            GetSurfaceProperty = m_SurfaceRenderingOriginY
        
        Case P2_SurfaceBlendUsingSRGBGamma
            GetSurfaceProperty = m_SurfaceBlendUsingSRGBGamma
    
    End Select

End Function

Public Sub SetSurfaceProperty(ByVal propID As PD_2D_SURFACE_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
    
        Case P2_SurfaceAntialiasing
            m_SurfaceAntialiasing = CLng(newSetting)
        
        Case P2_SurfacePixelOffset
            m_SurfacePixelOffset = CLng(newSetting)
            
        Case P2_SurfaceRenderingOriginX
            m_SurfaceRenderingOriginX = CSng(newSetting)
        
        Case P2_SurfaceRenderingOriginY
            m_SurfaceRenderingOriginY = CSng(newSetting)
            
        Case P2_SurfaceBlendUsingSRGBGamma
            m_SurfaceBlendUsingSRGBGamma = CBool(newSetting)
            
    End Select
    
    RelayPropertyChangeToBackend propID
    
End Sub

'Some property changes exist only at run-time.  Because these properties are not well-defined, they do not support
' serialization (e.g. you cannot save/load them to a string).  Clipping is one such property, because clipping may
' be undefined, a simple rect, or a complex region, and pd2D doesn't implement all those features... yet.
' (GDIPlus can retrieve region data as a byte array, so we could theoretically convert that to Base64 and store it
'  inside an XML tag.)
Public Function GetSurfaceClippingRegion(ByRef dstRegion As pd2DRegion) As Boolean

    If (m_SurfaceHandle <> 0) Then
        
        If (dstRegion Is Nothing) Then Set dstRegion = New pd2DRegion
        
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                Dim tmpHandle As Long
                tmpHandle = GDI_Plus.GDIPlus_GraphicsGetClipRegion(m_SurfaceHandle)
                dstRegion.AssignExternalHandleDirectly tmpHandle
                GetSurfaceClippingRegion = CBool(tmpHandle <> 0)
        End Select
        
    Else
        GetSurfaceClippingRegion = False
    End If
        
End Function

Public Function SetSurfaceClip_FromRectangle(ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean

    If (m_SurfaceHandle <> 0) Then
    
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                SetSurfaceClip_FromRectangle = GDI_Plus.GDIPlus_GraphicsSetClipRect(m_SurfaceHandle, rectLeft, rectTop, rectWidth, rectHeight, useCombineMode)
        End Select
        
    Else
        SetSurfaceClip_FromRectangle = False
    End If
        
End Function

Friend Function SetSurfaceClip_FromRectF(ByRef srcRectF As RECTF, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    With srcRectF
        SetSurfaceClip_FromRectF = SetSurfaceClip_FromRectangle(.Left, .Top, .Width, .Height)
    End With
End Function

Friend Function SetSurfaceClip_FromRectL(ByRef srcRectL As RECTL, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    With srcRectL
        SetSurfaceClip_FromRectL = SetSurfaceClip_FromRectangle(.Left, .Top, .Right - .Left, .Bottom - .Top)
    End With
End Function

Friend Function SetSurfaceClip_FromRegion(ByRef srcRegion As pd2DRegion, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean

    If (m_SurfaceHandle <> 0) Then
    
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                SetSurfaceClip_FromRegion = GDI_Plus.GDIPlus_GraphicsSetClipRegion(m_SurfaceHandle, srcRegion.GetHandle, useCombineMode)
        End Select
        
    Else
        SetSurfaceClip_FromRegion = False
    End If
    
End Function

Public Function SetSurfaceClip_None() As Boolean

    If (m_SurfaceHandle <> 0) Then
    
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                SetSurfaceClip_None = GDI_Plus.GDIPlus_GraphicsResetClipRegion(m_SurfaceHandle)
        End Select
        
    Else
        SetSurfaceClip_None = False
    End If
    
End Function

'Create a blank in-memory surface.  pd2DSurface will automatically manage the memory for this surface.
Public Function CreateBlankSurface(ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, Optional ByVal surfaceHasAlphaChannel As Boolean = True, Optional ByVal surfaceBackColor As Long = vbWhite, Optional ByVal surfaceOpacity As Single = 100#) As Boolean
    If (surfaceWidth <> 0) And (surfaceHeight <> 0) Then
        If (m_SurfaceDIB Is Nothing) Then Set m_SurfaceDIB = New pdDIB
        Dim targetColorDepth As Long
        If surfaceHasAlphaChannel Then targetColorDepth = 32 Else targetColorDepth = 24
        CreateBlankSurface = m_SurfaceDIB.CreateBlank(surfaceWidth, surfaceHeight, targetColorDepth, surfaceBackColor, surfaceOpacity * 2.55)
        If CreateBlankSurface Then m_SurfaceType = P2_ST_Bitmap
    Else
        InternalError "Invalid surface width/height", "When creating a blank surface, the width and height must be non-zero."
    End If
End Function

'Wrap this surface around some existing hDC.  It's fine to call this function multiple times on the same DC - changes to the handle
' are detected automatically
Public Function WrapSurfaceAroundDC(ByVal srcDC As Long) As Boolean
    
    If (srcDC = m_SurfaceDC) Then
        WrapSurfaceAroundDC = True
    Else
        
        If (srcDC <> 0) Then
        
            If (m_SurfaceHandle <> 0) Then ReleaseSurface
            
            'Call the backend-specific surface creation function
            Select Case m_SurfaceBackend
                
                Case P2_DefaultBackend, P2_GDIPlusBackend
                
                    m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(srcDC, IIf(m_SurfaceAntialiasing >= P2_AA_HighQuality, GP_SM_AntiAlias, GP_SM_None), IIf(m_SurfacePixelOffset >= P2_PO_Half, GP_POM_Half, GP_POM_None))
                    WrapSurfaceAroundDC = (m_SurfaceHandle <> 0)
                    
                    If WrapSurfaceAroundDC Then
                        m_SurfaceDC = srcDC
                        m_SurfaceType = P2_ST_WrapperOnly
                    End If
                    
            End Select
            
            'When debug mode is active, all object creations are reported back to the master Drawing2D module
            If (WrapSurfaceAroundDC And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, True
            
        End If
        
    End If
    
End Function

Public Function GetHandle() As Long
    GetHandle = m_SurfaceHandle
End Function

Friend Function GetSurfaceBackend() As PD_2D_RENDERING_BACKEND
    GetSurfaceBackend = m_SurfaceBackend
End Function

'I'm still debating whether it's wise to expose this function externally.  GDI+ handles DCs in weird ways; as long as it is only ever
' wrapped around *existing* DCs, you can intermix GDI+ and GDI calls without issue, and everything auto-synchronizes.  However, if you
' create a native GDI+ surface then use GdipGetDC() to retrieve a DC for it, you *must* call GdipReleaseDC when you're done (which is
' very unintuitive, given the way DCs usually work for VB programmers).
'
'As such, my current inclination is to always keep a GDI-backed copy of any surfaces created by this class.  This allows the caller to
' freely intermix bare GDI calls (like BitBlt) and pd2D drawing calls without consequence, and it also circumvents some really obnoxious
' GDI+ limitations (like being forced to keep a file around for the lifetime of an image object, if the image object is created from
' that file).
'
'Anyway, I mention this here as an FYI, but tl;dr: for now, this function is safe to use.
Public Function GetSurfaceDC() As Long
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        GetSurfaceDC = m_SurfaceDC
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceDC = m_SurfaceDIB.GetDIBDC Else GetSurfaceDC = 0
    Else
        GetSurfaceDC = 0
    End If
End Function

Public Function GetSurfaceAlphaSupport() As Boolean
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
        GetSurfaceAlphaSupport = CBool(m_SurfaceDCBitmap.BitsPerPixel = 32)
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceAlphaSupport = CBool(m_SurfaceDIB.GetDIBColorDepth = 32)
    Else
        GetSurfaceAlphaSupport = False
    End If
End Function

Public Function GetSurfaceHeight() As Long
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
        GetSurfaceHeight = m_SurfaceDCBitmap.Height
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceHeight = m_SurfaceDIB.GetDIBHeight
    Else
        GetSurfaceHeight = 0
    End If
End Function

Public Function GetSurfaceStride() As Long
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
        GetSurfaceStride = m_SurfaceDCBitmap.WidthBytes
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceStride = m_SurfaceDIB.GetDIBStride
    Else
        GetSurfaceStride = 0
    End If
End Function

Public Function GetSurfaceWidth() As Long
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
        GetSurfaceWidth = m_SurfaceDCBitmap.Width
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceWidth = m_SurfaceDIB.GetDIBWidth
    Else
        GetSurfaceWidth = 0
    End If
End Function

Public Function HasSurface() As Boolean
    HasSurface = CBool(m_SurfaceHandle <> 0)
End Function

'You only need to manually initialize a surface backend if you plan on using a non-default backend.  (Surfaces with the
' default backend are initialized automatically.)
Friend Function InitializeSurfaceBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_SurfaceBackend = targetBackend
    InitializeSurfaceBackend = True
End Function

Public Function ReleaseSurface() As Boolean
    
    If (m_SurfaceHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_SurfaceBackend
        
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleaseSurface = GDI_Plus.ReleaseGDIPlusGraphics(m_SurfaceHandle)
        
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseSurface Then
            m_SurfaceHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, False
            
            'Also reset any surface-type-specific values
            If (m_SurfaceType = P2_ST_WrapperOnly) Then
                m_SurfaceDC = 0
            ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
                Set m_SurfaceDIB = Nothing
            End If
            
            m_SurfaceType = P2_ST_Uninitialized
            
        End If
    
    Else
        ReleaseSurface = True
    End If
    
    'Releasing a handle also releases any underlying trackers (hDCs, pointers, etc)
    m_SurfaceDC = 0

End Function

Private Sub RelayPropertyChangeToBackend(ByVal propID As PD_2D_SURFACE_SETTINGS)
    
    'Property changes only need to be relayed if we've already created a handle
    If (m_SurfaceHandle <> 0) Then
    
        'Relays obviously vary by the active backend!
        Select Case m_SurfaceBackend
            
            Case P2_DefaultBackend, P2_GDIPlusBackend
                
                Select Case propID
                    Case P2_SurfaceAntialiasing
                        If (m_SurfaceAntialiasing >= P2_AA_HighQuality) Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_SM_AntiAlias
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_SM_None
                        End If
                        
                    Case P2_SurfacePixelOffset
                        If (m_SurfacePixelOffset >= P2_PO_Half) Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_POM_Half
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_POM_None
                        End If
                    
                    Case P2_SurfaceRenderingOriginX
                        GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, m_SurfaceRenderingOriginX
                        
                    Case P2_SurfaceRenderingOriginY
                        GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, m_SurfaceRenderingOriginY
                        
                    Case P2_SurfaceBlendUsingSRGBGamma
                        If m_SurfaceBlendUsingSRGBGamma Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceBlendUsingSRGBGamma, GP_CQ_GammaCorrected
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceBlendUsingSRGBGamma, GP_CQ_AssumeLinear
                        End If
                        
                End Select
                
            'In the future, other backends can be handled here...
            Case Else
                
        End Select
    
    End If
    
End Sub

Public Sub ResetAllProperties()
    Me.SetSurfaceProperty P2_SurfaceAntialiasing, P2_AA_None
    Me.SetSurfaceProperty P2_SurfacePixelOffset, P2_PO_Normal
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginX, 0#
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginY, 0#
    Me.SetSurfaceProperty P2_SurfaceBlendUsingSRGBGamma, False
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    m_SurfaceType = P2_ST_Uninitialized
    Me.ResetAllProperties
End Sub

Private Sub Class_Terminate()
    Me.ReleaseSurface
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DSurface"
End Sub
