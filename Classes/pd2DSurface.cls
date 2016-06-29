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
' 2) Bitmap (raster) surfaces.  The creation and destruction of such surfaces is handled internally, using a pddib object.
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

'The current surface handle.  Will be 0 if the surface is uninitialized.  The meaning of this value varies by backend.
Private m_SurfaceHandle As Long

'GDI+ uses a number of different image handle types.  A "Graphics" handle is like a DC.  It's an abstract value into
' which other, concrete image types can be select.  The above "SurfaceHandle" is actually a "Graphics" handle in disguise.
' Some GDI operations require an "Image" handle.  This handle may a bitmap or metafile.  (And if loaded from file, it may
' be even more abstract, like an EXIF file.)  In PD, it is almost certainly a bitmap object.  This handle is created
' on-demand only as necessary.  Any time the surface handle, above, is cleared, this value must also be cleared.
Private m_GdipImageHandle As Long

'If this surface was created by wrapping an existing hDC, this value will be non-zero.
Private m_SurfaceDC As Long

'If you need to query a wrapped DC for bitmap information (e.g. width/height), use this temporary object to do so.
Private m_SurfaceDCBitmap As GDI_Bitmap

'If this surface was created as a new surface, this DIB will serve as the surface backer.  (Similarly, the surface DC will simply
' wrap this DIB's hDC.)
Private m_SurfaceDIB As pdDIB

'This class internally manages what kind of surface it is (e.g. a lightweight wrapper around an existing hDC,
' or a fully-managed DIB surface).  This value is exposed via the GetSurfaceType function, but users should never
' have to mess with this; we only use it internally, in places like the Painter class, to figure out the best
' GDI+ interop method.  (Wrapped hDCs can't be used as 32-bpp surfaces, because alpha data is unreliable.)
Private m_SurfaceType As PD_2D_SurfaceType

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Surface settings that support get/set operations
Private m_SurfaceAntialiasing As PD_2D_Antialiasing
Private m_SurfaceBlendUsingSRGBGamma As Boolean
Private m_SurfacePixelOffset As PD_2D_PixelOffset
Private m_SurfaceRenderingOriginX As Single, m_SurfaceRenderingOriginY As Single
Private m_SurfaceResizeQuality As PD_2D_ResizeQuality

'Get/set individual settings.  Note that these just wrap the generic Get/SetPenProperty functions, below.
Public Function GetSurfaceAntialiasing() As PD_2D_Antialiasing
    GetSurfaceAntialiasing = Me.GetSurfaceProperty(P2_SurfaceAntialiasing)
End Function

Public Function GetSurfaceBlendUsingSRGBGamma() As Boolean
    GetSurfaceBlendUsingSRGBGamma = Me.GetSurfaceProperty(P2_SurfaceBlendUsingSRGBGamma)
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

Public Function GetSurfaceResizeQuality() As PD_2D_ResizeQuality
    GetSurfaceResizeQuality = Me.GetSurfaceProperty(P2_SurfaceResizeQuality)
End Function

Public Sub SetSurfaceAntialiasing(ByVal newSetting As PD_2D_Antialiasing)
    Me.SetSurfaceProperty P2_SurfaceAntialiasing, newSetting
End Sub

Public Sub SetSurfaceBlendUsingSRGBGamma(ByVal newSetting As Boolean)
    Me.SetSurfaceProperty P2_SurfaceBlendUsingSRGBGamma, newSetting
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

Public Sub SetSurfaceResizeQuality(ByVal newSetting As PD_2D_ResizeQuality)
    Me.SetSurfaceProperty P2_SurfaceResizeQuality, newSetting
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
            
        Case P2_SurfaceResizeQuality
            GetSurfaceProperty = m_SurfaceResizeQuality
    
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
            
        Case P2_SurfaceResizeQuality
            m_SurfaceResizeQuality = CLng(newSetting)
            
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
        
        If (m_SurfaceHandle <> 0) Then ReleaseSurface
    
        'Prep a GDI DIB for use as our blank surface
        If (m_SurfaceDIB Is Nothing) Then Set m_SurfaceDIB = New pdDIB
        Dim targetColorDepth As Long
        If surfaceHasAlphaChannel Then targetColorDepth = 32 Else targetColorDepth = 24
        CreateBlankSurface = m_SurfaceDIB.CreateBlank(surfaceWidth, surfaceHeight, targetColorDepth, surfaceBackColor, surfaceOpacity * 2.55)
        
        If CreateBlankSurface Then
        
            m_SurfaceType = P2_ST_Bitmap
            
            'Finally, wrap the appropriate backend around this newly created DC
            Select Case m_SurfaceBackend
                
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(m_SurfaceDIB.GetDIBDC, IIf(m_SurfaceAntialiasing >= P2_AA_HighQuality, GP_SM_AntiAlias, GP_SM_None), IIf(m_SurfacePixelOffset >= P2_PO_Half, GP_POM_Half, GP_POM_None))
                    CreateBlankSurface = (m_SurfaceHandle <> 0)
                    
            End Select
            
        End If
    Else
        InternalError "Invalid surface width/height", "When creating a blank surface, the width and height must be non-zero."
    End If
    
End Function

'Wrap this surface around some existing hDC.  It's fine to call this function multiple times on the same DC - changes to the handle
' are detected automatically.
Public Function WrapSurfaceAroundDC(ByVal srcDC As Long) As Boolean
    
    If (srcDC <> 0) Then
    
        If (srcDC = m_SurfaceDC) Then
            WrapSurfaceAroundDC = True
        Else
            
            If (m_SurfaceHandle <> 0) Then ReleaseSurface
            
            'Call the backend-specific surface creation function
            Select Case m_SurfaceBackend
                
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    
                    m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(srcDC, IIf(m_SurfaceAntialiasing >= P2_AA_HighQuality, GP_SM_AntiAlias, GP_SM_None), IIf(m_SurfacePixelOffset >= P2_PO_Half, GP_POM_Half, GP_POM_None))
                    WrapSurfaceAroundDC = (m_SurfaceHandle <> 0)
                    
                    If WrapSurfaceAroundDC Then
                        m_SurfaceDC = srcDC
                        m_SurfaceType = P2_ST_WrapperOnly
                        
                        'Apply any other stored settings to the newly created surface
                        RelayPropertyChangeToBackend P2_SurfaceBlendUsingSRGBGamma
                        RelayPropertyChangeToBackend P2_SurfaceRenderingOriginX
                        RelayPropertyChangeToBackend P2_SurfaceRenderingOriginY
                        RelayPropertyChangeToBackend P2_SurfaceResizeQuality
                        
                    Else
                        InternalError "GDI+ DC wrapper failed", "GDI+ failed to wrap the hDC passed to pd2DSurface.WrapSurfaceAroundDC"
                    End If
                    
            End Select
            
            'When debug mode is active, all object creations are reported back to the master Drawing2D module
            If (WrapSurfaceAroundDC And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, True
            
        End If
    
    Else
        InternalError "hDC is null", "pd2DSurface.WrapSurfaceAroundDC doesn't work on null DCs."
    End If
    
End Function

Public Function GetHandle() As Long
    GetHandle = m_SurfaceHandle
End Function

'Retrieving an Image-type handle is an expensive operation for graphics objects wrapped around a DC.  Whenever possible,
' restrict usage of this function to internally-managed DIBs (which are very fast by comparison).
Friend Function GetGdipImageHandle() As Long
    
    If (m_GdipImageHandle = 0) Then
        
        'If this surface is just a wrapper around a DC, we need to create a temporary copy of it, and wrap *that* instead.
        ' AFAIK, this is the fastest, most reliable way to generate bitmap data from an arbitrary DC (which may not contain
        ' bitmap data at all).
        
        'Note that alpha data *will* be ignored, by design.  The only way to fix this would be to ask the caller for
        ' permission to use alpha bytes, which may not be valid as many (most?) DCs will have DDBs selected into them,
        ' meaning alpha bytes are device-dependent and not necessarily relevant or correct.
        If (m_SurfaceType = P2_ST_WrapperOnly) Then
            
            'Allocating memory for new DIBs is expensive, so whenever possible, skip these steps
            If (m_SurfaceDIB Is Nothing) Then Set m_SurfaceDIB = New pdDIB
            If (m_SurfaceDIB.GetDIBWidth <> Me.GetSurfaceWidth) Or (m_SurfaceDIB.GetDIBHeight <> Me.GetSurfaceHeight) Then m_SurfaceDIB.CreateBlank Me.GetSurfaceWidth, Me.GetSurfaceHeight, 24, 0
            
            'BitBlt lets us quickly copy the source DC's contents, regardless of the source DC's selected object
            GDI.BitBltWrapper m_SurfaceDIB.GetDIBDC, 0, 0, m_SurfaceDIB.GetDIBWidth, m_SurfaceDIB.GetDIBHeight, m_SurfaceDC, 0, 0, vbSrcCopy
            
        End If
        
        If (Not GDI_Plus.GetGdipBitmapHandleFromDIB(m_GdipImageHandle, m_SurfaceDIB)) Then InternalError "GetGdipImageHandle failed", "Could not create a GDI+ Image (bitmap) object from wrapped hDC."
        
    'If we already have a handle, but we are in "wrapper" mode, we need to make sure our duplicate DIB copy is still up-to-date.
    Else
    
        If (m_SurfaceType = P2_ST_WrapperOnly) Then
        
            If (m_SurfaceDIB.GetDIBWidth <> Me.GetSurfaceWidth) Or (m_SurfaceDIB.GetDIBHeight <> Me.GetSurfaceHeight) Then
                ReleaseExtraInternalObjects
                m_SurfaceDIB.CreateBlank Me.GetSurfaceWidth, Me.GetSurfaceHeight, 24, 0
            End If
            
            GDI.BitBltWrapper m_SurfaceDIB.GetDIBDC, 0, 0, m_SurfaceDIB.GetDIBWidth, m_SurfaceDIB.GetDIBHeight, m_SurfaceDC, 0, 0, vbSrcCopy
            
            If (m_GdipImageHandle = 0) Then
                If Not GDI_Plus.GetGdipBitmapHandleFromDIB(m_GdipImageHandle, m_SurfaceDIB) Then InternalError "GetGdipImageHandle failed", "Could not create a GDI+ Image (bitmap) object from wrapped hDC."
            End If
            
        End If
        
    End If
    
    GetGdipImageHandle = m_GdipImageHandle
    
End Function

Friend Function GetSurfaceBackend() As PD_2D_RENDERING_BACKEND
    GetSurfaceBackend = m_SurfaceBackend
End Function

'Users never need to call this function; it is only used to accelerate various drawing options
Friend Function GetSurfaceType() As PD_2D_SurfaceType
    GetSurfaceType = m_SurfaceType
End Function

'I'm still debating whether it's wise to expose this function externally.  GDI+ handles DCs in weird ways;
' as long as it is only ever wrapped around *existing* DCs, you can usually intermix GDI+ and GDI calls
' without running into problems.  However, if you create a native GDI+ surface then use GdipGetDC() to
' retrieve a DC for it, you *must* call GdipReleaseDC when you're done (which is very unintuitive, given the
' way DCs usually work for VB programmers).
'
'As such, my current inclination is to always keep a GDI-backed copy of any surfaces created by this class.
' This simplifies situations where the wants to intermix bare GDI calls (like BitBlt) and pd2D drawing calls,
' and it also circumvents some obnoxious GDI+ limitations (like being forced to keep a source file around for
' the lifetime of an image object created from that file).
'
'Anyway, I mention this here as an FYI, but tl;dr: for now, this function is safe to use, and you don't have
' to free the DC after an operation completes.
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

'Erase the current surface's contents.  Roughly equivalent to .Cls in VB, except that this function accepts a
' color and/or alpha to use when erasing.
'
'Note that wrapped GDI surfaces will *not* touch alpha, by design, to avoid screwing up the target DC's behavior.
Public Sub EraseSurfaceContents(Optional ByVal newColor As Long = vbWhite, Optional ByVal newAlpha As Single = 0#)
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        GDI.FillRectToDC m_SurfaceDC, 0, 0, Me.GetSurfaceWidth, Me.GetSurfaceHeight, newColor
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then m_SurfaceDIB.FillWithColor newColor, newAlpha
    End If
End Sub

Public Function ReleaseSurface() As Boolean
    
    'If any extra, specialized handles have been created, free them first.  (Some of these handles may wrap data
    ' we are about to free, so releasing them first prevents double-freeing crashes.)
    ReleaseExtraInternalObjects
    
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
                
                'Wrapped surfaces may use a temporary DIB for certain operations, so free that temp DIB if it exists
                If (Not m_SurfaceDIB Is Nothing) Then Set m_SurfaceDIB = Nothing
                
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

Private Sub ReleaseExtraInternalObjects()

    'GDI+ "Image" handles may be created on-demand, if required by a given paint function
    If (m_GdipImageHandle <> 0) Then
        GDI_Plus.ReleaseGDIPlusImage m_GdipImageHandle
        m_GdipImageHandle = 0
    End If
    
End Sub

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
                        
                    Case P2_SurfaceResizeQuality
                        Select Case m_SurfaceResizeQuality
                            Case P2_RQ_Fast
                                GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceResizeQuality, GP_IM_NearestNeighbor
                            Case P2_RQ_Bilinear
                                GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceResizeQuality, GP_IM_HighQualityBilinear
                            Case P2_RQ_Bicubic
                                GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceResizeQuality, GP_IM_HighQualityBicubic
                        End Select
                        
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
    Me.SetSurfaceProperty P2_SurfaceResizeQuality, P2_RQ_Fast
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

